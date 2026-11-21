#!/usr/bin/env python3
"""Compare a reconstructed PowerPoint to its source images and element manifest."""

import argparse
import base64
import hashlib
from io import BytesIO
import json
from math import isclose, isfinite
import os
from pathlib import Path
from posixpath import join, normpath
import re
import shutil
import subprocess
import tempfile
import unicodedata
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageOps
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE, XL_TICK_LABEL_POSITION
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.opc.constants import CONTENT_TYPE
from pptx.oxml.ns import qn


BACKGROUND_NAME = "__background__"
PIXEL_DELTA_THRESHOLD = 24
STRUCTURE_TOLERANCE_PX = 2.0
ROTATION_TOLERANCE_DEGREES = 0.25
OBJECT_OUTLINES = {"text": "#f97316", "image": "#06b6d4", "shape": "#a855f7", "chart": "#1d4ed8"}
DRAWING_TYPES = {MSO_SHAPE_TYPE.AUTO_SHAPE, MSO_SHAPE_TYPE.FREEFORM, MSO_SHAPE_TYPE.LINE}
NATIVE_GEOMETRIES = {
    "rect": MSO_SHAPE.RECTANGLE,
    "round_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
    "ellipse": MSO_SHAPE.OVAL,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
}
NATIVE_CHART_TYPES = {
    "bar": {XL_CHART_TYPE.BAR_CLUSTERED},
    "column": {XL_CHART_TYPE.COLUMN_CLUSTERED},
    "line": {XL_CHART_TYPE.LINE, XL_CHART_TYPE.LINE_MARKERS},
    "scatter": {XL_CHART_TYPE.XY_SCATTER, XL_CHART_TYPE.XY_SCATTER_LINES, XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS},
}
WORKSHEET_XML = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
WORKBOOK_REFERENCE = re.compile(
    r"(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]{1,3})\$?([1-9][0-9]*)"
    r"(?::\$?([A-Z]{1,3})\$?([1-9][0-9]*))?\Z",
    re.IGNORECASE,
)


def normalized_text(value):
    newlines = value.replace("\r\n", "\n").replace("\r", "\n").replace("\v", "\n")
    return unicodedata.normalize("NFC", newlines)


def shape_bounds(shape, slide_width, slide_height, width, height):
    values = (shape.left, shape.top, shape.width, shape.height)
    if any(value is None for value in values):
        return None
    x, y, w, h = values
    return [x * width / slide_width, y * height / slide_height,
            w * width / slide_width, h * height / slide_height]


def text_styles(shape):
    styles = []
    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            font = run.font
            properties = run._r.find(qn("a:rPr"))
            east_asian = None if properties is None else properties.find(qn("a:ea"))
            color = font.color
            style = {
                "font": font.name, "font_ea": None if east_asian is None else east_asian.get("typeface"),
                "color": f"#{color.rgb}" if color.type == MSO_COLOR_TYPE.RGB else None,
                "font_pt": font.size.pt if font.size else None,
                "bold": font.bold, "italic": font.italic,
            }
            if style not in styles:
                styles.append(style)
    return styles


def chart_value_matches(actual, expected):
    if isinstance(expected, str):
        return isinstance(actual, str) and normalized_text(actual) == normalized_text(expected)
    # DrawingML and Excel serialize floating point values differently. Keep a
    # small relative allowance for decimal round-trip, and never round to zero.
    return (type(actual) in (int, float) and isfinite(actual)
            and isclose(actual, expected, rel_tol=1e-14, abs_tol=0.0))


class ChartWorkbookValues:
    """Read literal chart-source cells from the actual embedded .xlsx part."""

    def __init__(self, xlsx_bytes):
        self.sheets = {}
        with ZipFile(BytesIO(xlsx_bytes)) as archive:
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            rels = list(ET.fromstring(archive.read("xl/_rels/workbook.xml.rels")))
            targets = {rel.get("Id"): rel for rel in rels if rel.get("TargetMode") != "External"}
            string_rel = next((rel for rel in targets.values()
                               if rel.get("Type", "").endswith("/sharedStrings")), None)
            strings = []
            if string_rel is not None:
                table = ET.fromstring(archive.read(self.part_name(string_rel)))
                strings = [self.rich_text(item) for item in table.findall("s:si", WORKSHEET_XML)]
            for sheet in workbook.findall("s:sheets/s:sheet", WORKSHEET_XML):
                sheet_rel = targets[sheet.get(qn("r:id"))]
                if not sheet_rel.get("Type", "").endswith("/worksheet"):
                    continue
                root = ET.fromstring(archive.read(self.part_name(sheet_rel)))
                cells = root.findall("s:sheetData/s:row/s:c", WORKSHEET_XML)
                self.sheets[sheet.get("name").casefold()] = {
                    cell.get("r"): self.cell_value(cell, strings) for cell in cells
                }

    @staticmethod
    def part_name(rel):
        target = rel.get("Target")
        return normpath(target.lstrip("/") if target.startswith("/") else join("xl", target))

    @staticmethod
    def rich_text(element):
        runs = element.findall("s:t", WORKSHEET_XML) + element.findall("s:r/s:t", WORKSHEET_XML)
        return "".join(run.text or "" for run in runs)

    @classmethod
    def cell_value(cls, cell, strings):
        if cell.find("s:f", WORKSHEET_XML) is not None:
            return None
        cell_type = cell.get("t", "n")
        value = cell.findtext("s:v", namespaces=WORKSHEET_XML)
        if cell_type == "s" and value is not None:
            index = int(value)
            return strings[index] if 0 <= index < len(strings) else None
        if cell_type == "inlineStr":
            inline = cell.find("s:is", WORKSHEET_XML)
            return cls.rich_text(inline) if inline is not None else None
        if cell_type == "str":
            return value
        if cell_type == "n" and value is not None:
            return float(value)
        return None

    def range_values(self, formula, size):
        match = WORKBOOK_REFERENCE.fullmatch(formula)
        if not match:
            raise ValueError(f"unsupported workbook reference {formula!r}")
        quoted_sheet, sheet, column, first_row, end_column, end_row = match.groups()
        cells = self.sheets.get((quoted_sheet.replace("''", "'") if quoted_sheet is not None else sheet).casefold())
        first = int(first_row)
        last = int(end_row) if end_row is not None else first
        if cells is None or (end_column is not None and column.upper() != end_column.upper()) or last - first + 1 != size:
            raise ValueError(f"workbook reference {formula!r} does not cover the expected column range")
        return [cells.get(f"{column.upper()}{row}") for row in range(first, last + 1)]


def audit_chart_range(series, path, expected, workbook, issues, label, workbook_issues):
    references = series.xpath(path)
    if len(references) != 1:
        issues.append(f"Chart {label} has no single cached, workbook-backed reference.")
        workbook_issues.append(label)
        return None
    reference = references[0]
    cache_tag = "c:strCache" if isinstance(expected[0], str) else "c:numCache"
    cache = reference.find(qn(cache_tag))
    actual = None
    if cache is not None:
        count = cache.find(qn("c:ptCount"))
        points = cache.findall(qn("c:pt"))
        if count is not None and int(count.get("val")) == len(expected) == len(points):
            values = {int(point.get("idx")): point.findtext(qn("c:v")) for point in points}
            if set(values) == set(range(len(expected))) and all(value is not None for value in values.values()):
                actual = [values[i] for i in range(len(expected))]
                if cache_tag == "c:numCache":
                    actual = [float(value) for value in actual]
                    if not all(isfinite(value) for value in actual):
                        actual = None
    if actual is None or not all(chart_value_matches(a, e) for a, e in zip(actual, expected)):
        issues.append(f"Chart {label} cached data differ from the manifest.")

    formula = reference.findtext(qn("c:f"))
    if workbook is not None:
        try:
            values = workbook.range_values(formula or "", len(expected))
        except ValueError as error:
            issues.append(f"Chart {label} has an invalid Excel workbook reference: {error}.")
            workbook_issues.append(label)
        else:
            if not all(chart_value_matches(a, e) for a, e in zip(values, expected)):
                issues.append(f"Chart {label} Excel workbook data differ from the manifest.")
                workbook_issues.append(label)
    return actual


def audit_chart_axis(axis, options, issues, description):
    for key, actual in (
        ("min", axis.minimum_scale), ("max", axis.maximum_scale), ("major_unit", axis.major_unit),
        ("number_format", axis.tick_labels.number_format),
    ):
        if key in options and not chart_value_matches(actual, options[key]):
            issues.append(f"Chart {description} {key} differs from the manifest.")
    for key, actual in (
        ("visible", axis.visible), ("show_major_gridlines", axis.has_major_gridlines),
        ("show_tick_labels", axis.tick_label_position != XL_TICK_LABEL_POSITION.NONE),
    ):
        if key in options and actual != options[key]:
            issues.append(f"Chart {description} {key} differs from the manifest.")
    if "number_format" in options and axis.tick_labels.number_format_is_linked:
        issues.append(f"Chart {description} unexpectedly uses the workbook's linked number format.")
    if "title" in options and (
        not axis.has_title or not chart_value_matches(axis.axis_title.text_frame.text, options["title"])
    ):
        issues.append(f"Chart {description} title differs from the manifest.")


def check_chart(shape, element, result):
    issues = result["issues"]
    result["chart_workbook_status"] = "not_verified"
    result["chart_source_note"] = element.get("source_note")
    if not shape.has_chart:
        issues.append("Expected a native editable PowerPoint chart, not a flattened image or drawing.")
        return
    try:
        chart = shape.chart
        kind = chart.chart_type
        result["actual_chart_type"] = kind.name
        if kind not in NATIVE_CHART_TYPES.get(element["chart_type"], set()) or len(chart.plots) != 1:
            issues.append("PowerPoint chart type or plot count differs from the manifest.")
        if "title" in element and (
            not chart.has_title or not chart_value_matches(chart.chart_title.text_frame.text, element["title"])
        ):
            issues.append("Chart title differs from the manifest.")
        if "source_note" in element:
            description = shape._element.nvGraphicFramePr.cNvPr.get("descr")
            if not chart_value_matches(description, element["source_note"]):
                issues.append("Chart description does not preserve the source note.")
        if "show_legend" in element and chart.has_legend != element["show_legend"]:
            issues.append("Chart legend visibility differs from the manifest.")
        if "data_labels" in element and chart.plots[0].has_data_labels != element["data_labels"]:
            issues.append("Chart data labels visibility differs from the manifest.")

        workbook = None
        workbook_issues = []
        try:
            part = chart.part.chart_workbook.xlsx_part
            if part is None or part.content_type != CONTENT_TYPE.SML_SHEET:
                issues.append("Chart has no embedded editable Excel workbook.")
            else:
                workbook = ChartWorkbookValues(part.blob)
        except (AttributeError, BadZipFile, ET.ParseError, IndexError, KeyError, OSError, TypeError, ValueError):
            issues.append("Chart embedded editable Excel workbook cannot be read.")

        actual_series = chart.plots[0]._element.findall(qn("c:ser"))
        expected_series = element["series"]
        if len(actual_series) != len(expected_series):
            issues.append("Chart series count differs from the manifest.")
            workbook_issues.append("series count")
        is_scatter = element["chart_type"] == "scatter"
        recorded_series = []
        for index, (actual, expected) in enumerate(zip(actual_series, expected_series), 1):
            label = f"series {index}"
            order = actual.find(qn("c:order"))
            if order is None or order.get("val") != str(index - 1):
                issues.append(f"Chart {label} order differs from the manifest.")
            fields = [
                ("name", "c:tx/c:strRef", [expected["name"]]),
                ("values", "c:yVal/c:numRef" if is_scatter else "c:val/c:numRef", expected["values"]),
                ("x_values", "c:xVal/c:numRef", expected["x_values"]) if is_scatter
                else ("categories", "c:cat/c:strRef", element["categories"]),
            ]
            row = {
                key: audit_chart_range(actual, path, values, workbook, issues, f"{label} {key}", workbook_issues)
                for key, path, values in fields
            }
            if row["name"] is not None:
                row["name"] = row["name"][0]
            recorded_series.append(row)
        result["actual_chart_series"] = recorded_series
        result["chart_workbook_status"] = "verified" if workbook is not None and not workbook_issues else "invalid"

        if element.get("value_axis"):
            audit_chart_axis(chart.value_axis, element["value_axis"], issues, "value axis")
        if is_scatter and element.get("x_axis"):
            audit_chart_axis(chart.category_axis, element["x_axis"], issues, "X axis")
        if "category_axis_title" in element:
            axis = chart.category_axis
            if not axis.has_title or not chart_value_matches(
                axis.axis_title.text_frame.text, element["category_axis_title"]
            ):
                issues.append("Chart category axis title differs from the manifest.")
    except (AttributeError, IndexError, KeyError, NotImplementedError, OverflowError, TypeError, ValueError):
        issues.append("PowerPoint chart data or formatting cannot be read.")


def check_object(element, matching, presentation, width, height, index):
    object_id = element["id"]
    kind = element["kind"]
    expected_box = element["box"]
    result = {"id": object_id, "kind": kind, "expected_box_px": expected_box, "issues": []}
    issues = result["issues"]
    if len(matching) != 1:
        issues.append(f"Expected one object named {object_id!r}; found {len(matching)}.")
        return result

    shape = matching[0]
    result["actual_type"] = shape.shape_type.name
    result["slide_shape_index"] = index[object_id]
    if kind == "image" and shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        issues.append("Expected an embedded image.")
    if kind == "shape" and shape.shape_type not in DRAWING_TYPES:
        issues.append("Expected a native PowerPoint drawing object.")
    if kind == "shape" and shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
        result["actual_geometry"] = shape.auto_shape_type.name
        expected_geometry = NATIVE_GEOMETRIES.get(element.get("shape"))
        if expected_geometry is not None and shape.auto_shape_type != expected_geometry:
            issues.append("Native PowerPoint shape geometry differs from the manifest.")
    if kind == "chart":
        check_chart(shape, element, result)
    if kind == "text":
        if not shape.has_text_frame:
            issues.append("Expected a native PowerPoint text frame.")
        else:
            expected_text = element.get("text", "")
            result["actual_text"] = shape.text
            if normalized_text(shape.text) != normalized_text(expected_text):
                result["expected_text"] = expected_text
                issues.append("PPT text differs from the manifest.")
            styles = text_styles(shape)
            result["actual_styles"] = styles
            for key in ("font", "font_ea", "color", "bold", "italic"):
                if key in element and (
                    not styles or any(str(style[key]).casefold() != str(element[key]).casefold() for style in styles)
                ):
                    issues.append(f"PPT text style for {key} differs from the manifest.")
            points_per_pixel = presentation.slide_width / (width * 12700)
            if "font_px" in element:
                expected_points = max(1, min(4000, element["font_px"] * points_per_pixel))
                if not styles or any(style["font_pt"] is None or
                                     abs(style["font_pt"] - expected_points) > 0.02 for style in styles):
                    issues.append("PPT text size differs from the manifest.")
            if "line_height_px" in element:
                expected_points = element["line_height_px"] * points_per_pixel
                for paragraph in shape.text_frame.paragraphs:
                    spacing = paragraph.line_spacing
                    if not hasattr(spacing, "pt") or abs(spacing.pt - expected_points) > 0.02:
                        issues.append("PPT text line height differs from the manifest's fixed spacing.")
                        break

    bounds = shape_bounds(shape, presentation.slide_width, presentation.slide_height, width, height)
    if bounds is None:
        issues.append("PPT object has no position or dimensions.")
    else:
        result["actual_box_px"] = [round(value, 3) for value in bounds]
        deltas = [actual - expected for actual, expected in zip(bounds, expected_box)]
        result["box_delta_px"] = [round(value, 3) for value in deltas]
        if max(abs(delta) for delta in deltas) > STRUCTURE_TOLERANCE_PX:
            issues.append("PPT position or dimensions exceed tolerance.")

    angle = float(shape.rotation or 0)
    expected_angle = float(element.get("rotation", 0))
    result["actual_rotation_degrees"] = angle
    if abs((angle - expected_angle + 180) % 360 - 180) > ROTATION_TOLERANCE_DEGREES:
        issues.append("PPT rotation exceeds tolerance.")
    return result


def check_slide(presentation, index, slide):
    width, height = slide["width"], slide["height"]
    report = {"id": slide["id"], "slide_number": index + 1, "issues": [], "objects": []}
    if index >= len(presentation.slides):
        report["issues"].append("No corresponding PowerPoint slide.")
        return report

    shapes = list(presentation.slides[index].shapes)
    by_name = {}
    z_index = {}
    for shape_index, shape in enumerate(shapes):
        by_name.setdefault(shape.name, []).append(shape)
        z_index[shape.name] = shape_index

    expected_order = [element["id"] for element in slide["elements"]]
    element_ids = set(expected_order)
    mapped_order = [shape.name for shape in shapes if shape.name in element_ids]
    if len(mapped_order) == len(expected_order) and mapped_order != expected_order:
        report["issues"].append("PPT element order differs from the manifest's bottom-to-top order.")
    unmapped = [shape.name for shape in shapes if shape.name not in element_ids | {BACKGROUND_NAME}]
    if unmapped:
        report["unmapped_shapes"] = unmapped
        report["issues"].append("PPT includes shapes without manifest IDs.")

    backgrounds = by_name.get(BACKGROUND_NAME, [])
    if not backgrounds:
        report["issues"].append("The reconstructed background object is missing.")
    elif len(backgrounds) > 1:
        report["issues"].append("PPT has more than one background object.")
    elif backgrounds:
        background = backgrounds[0]
        if background.shape_type != MSO_SHAPE_TYPE.PICTURE:
            report["issues"].append("The reconstructed background is not an embedded picture.")
        expected = [0, 0, width, height]
        actual = shape_bounds(background, presentation.slide_width, presentation.slide_height, width, height)
        if not actual or max(abs(a - b) for a, b in zip(actual, expected)) > STRUCTURE_TOLERANCE_PX:
            report["issues"].append("PPT background does not cover the source slide bounds.")
        if shapes[0].name != BACKGROUND_NAME:
            report["issues"].append("PPT background is not behind every other object.")

    expected_aspect = width / height
    actual_aspect = presentation.slide_width / presentation.slide_height
    if abs(expected_aspect / actual_aspect - 1) > 0.002:
        report["issues"].append("PPT slide and source have different aspect ratios.")

    for element in slide["elements"]:
        matching = by_name.get(element["id"], [])
        report["objects"].append(check_object(element, matching, presentation, width, height, z_index))
    return report


POWERSHELL_RENDER = r"""
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$config = ConvertFrom-Json -InputObject ([IO.File]::ReadAllText($env:CODEX_IMAGE_PPT_QA_RENDER_CONFIG))
$application = $null
$deck = $null
try {
    $application = New-Object -ComObject PowerPoint.Application
    $deck = $application.Presentations.Open($config.deck, -1, 0, 0)
    foreach ($slide in $config.slides) {
        if ([int]$slide.number -le $deck.Slides.Count) {
            $page = $deck.Slides.Item([int]$slide.number)
            try { $page.Export($slide.path, 'PNG', [int]$slide.width, [int]$slide.height) }
            finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($page) }
        }
    }
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
finally {
    if ($null -ne $deck) {
        $deck.Close()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($deck)
    }
    if ($null -ne $application) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($application)
    }
}
"""


def render_slides(pptx_path, output, slides, renderer):
    if renderer == "none":
        return {"status": "skipped", "reason": "Rendering disabled."}

    powershell = shutil.which("powershell.exe") if os.name == "nt" else None
    if not powershell:
        return {"status": "unavailable", "reason": "Windows PowerShell is not available."}

    encoded_script = base64.b64encode(POWERSHELL_RENDER.encode("utf-16-le")).decode("ascii")
    with tempfile.TemporaryDirectory(prefix="render-deck-", dir=output) as temporary:
        isolated_deck = Path(temporary, "verification-copy.pptx")
        shutil.copyfile(pptx_path, isolated_deck)
        render_config = Path(temporary, "render-config.json")
        render_config.write_text(
            json.dumps({"deck": str(isolated_deck), "slides": slides}, ensure_ascii=False),
            encoding="utf-8",
        )
        env = os.environ.copy()
        env["CODEX_IMAGE_PPT_QA_RENDER_CONFIG"] = str(render_config)
        try:
            result = subprocess.run(
                [powershell, "-NoProfile", "-NonInteractive", "-EncodedCommand", encoded_script],
                env=env, capture_output=True, encoding="utf-8", errors="replace", timeout=180,
                creationflags=subprocess.CREATE_NO_WINDOW, check=False,
            )
        except (OSError, subprocess.TimeoutExpired) as error:
            return {"status": "failed", "backend": "powerpoint", "reason": str(error)}
        if result.returncode:
            reason = (result.stderr.strip() or result.stdout.strip() or "Export failed")[:1200]
            return {"status": "failed", "backend": "powerpoint", "reason": reason}
    return {"status": "exported", "backend": "powerpoint"}


def outline_box(box, width, height):
    x, y, w, h = box
    return (
        max(0, min(width, int(np.floor(x)))),
        max(0, min(height, int(np.floor(y)))),
        max(0, min(width, int(np.ceil(x + w)))),
        max(0, min(height, int(np.ceil(y + h)))),
    )


def annotated_inventory(source, elements):
    annotated = source.copy()
    draw = ImageDraw.Draw(annotated)
    for order, element in enumerate(elements, 1):
        x, y, right, bottom = outline_box(element["box"], *source.size)
        color = OBJECT_OUTLINES.get(element["kind"], "#ef4444")
        draw.rectangle((x, y, right - 1, bottom - 1), outline=color, width=2)
        label = str(order)
        anchor_x = min(max(0, source.width - 32), x)
        anchor_y = min(max(0, source.height - 15), y)
        draw.rectangle((anchor_x, anchor_y, anchor_x + 9 * len(label) + 4, anchor_y + 15), fill="#111827")
        draw.text((anchor_x + 3, anchor_y + 1), label, fill=color)
    return annotated


def measure(delta):
    if not delta.size:
        return None
    salient = np.count_nonzero(delta > PIXEL_DELTA_THRESHOLD)
    return {
        "pixels": int(delta.size),
        "mean_max_channel_delta": round(float(np.mean(delta)), 3),
        "p95_max_channel_delta": round(float(np.quantile(delta, 0.95)), 3),
        "changed_fraction": round(salient / delta.size, 5),
    }


def localized_changes(delta, available):
    changed = (delta > PIXEL_DELTA_THRESHOLD) & available
    _, _, components, _ = cv2.connectedComponentsWithStats(changed.astype(np.uint8), connectivity=8)
    min_area = max(12, int(delta.size * 0.00001))
    findings = []
    for left, top, width, height, area in components[1:]:
        if area >= min_area:
            findings.append({
                "box": [int(left), int(top), int(width), int(height)],
                "changed_pixels": int(area),
            })
    findings.sort(key=lambda item: item["changed_pixels"], reverse=True)
    return {"count": len(findings), "largest": findings[:20]}


def compare_render(source, rendered, slide, prefix, output):
    width, height = source.size
    delta = np.max(np.abs(
        np.asarray(source, dtype=np.int16) - np.asarray(rendered, dtype=np.int16)
    ), axis=2)
    owned_pixels = np.zeros((height, width), dtype=bool)
    text_pixels = np.zeros_like(owned_pixels)
    chart_pixels = np.zeros_like(owned_pixels)
    regions = []
    review_flags = []
    for number, element in enumerate(slide["elements"], 1):
        x, y, right, bottom = outline_box(element["box"], width, height)
        owned_pixels[y:bottom, x:right] = True
        if element["kind"] == "text":
            text_pixels[y:bottom, x:right] = True
        region_stats = measure(delta[y:bottom, x:right])
        region = {"number": number, "id": element["id"], "stats": region_stats}
        regions.append(region)
        if element["kind"] == "chart":
            chart_pixels[y:bottom, x:right] = True
            ex, ey, er, eb = outline_box(element.get("erase", element["box"]), width, height)
            chart_pixels[ey:eb, ex:er] = True
            owned_pixels[ey:eb, ex:er] = True
            region["comparison_role"] = "chart_layout_reference_only"
            continue
        threshold = 0.15 if element["kind"] == "text" else 0.06
        if region_stats and region_stats["changed_fraction"] > threshold:
            review_flags.append({"region": element["id"], "reason": "Large visual difference inside the source box."})

    outside_text = measure(delta[~text_pixels])
    outside_text_and_charts = measure(delta[~(text_pixels | chart_pixels)])
    outside_objects = measure(delta[~owned_pixels])
    unassigned_changes = localized_changes(delta, ~owned_pixels)
    if outside_text_and_charts and outside_text_and_charts["changed_fraction"] > 0.02:
        if chart_pixels.any():
            review_flags.append({
                "region": "outside_text_and_charts",
                "reason": "Visual differences extend beyond source text and native chart boxes.",
            })
        else:
            review_flags.append({"region": "outside_text", "reason": "Visual differences extend beyond source text boxes."})
    if unassigned_changes["count"] or (outside_objects and outside_objects["changed_fraction"] > 0.01):
        review_flags.append({"region": "outside_objects", "reason": "Review visual changes outside inventoried elements."})

    tiles = []
    for row in range(5):
        for column in range(8):
            left, right = column * width // 8, (column + 1) * width // 8
            top, bottom = row * height // 5, (row + 1) * height // 5
            tile = delta[top:bottom, left:right]
            stats = measure(tile[~owned_pixels[top:bottom, left:right]])
            if stats and stats["changed_fraction"] > 0.05:
                tiles.append({"box": [left, top, right - left, bottom - top], "stats": stats})

    heatmap = np.zeros((height, width, 3), dtype=np.uint8)
    heatmap[:, :, 0] = np.minimum(255, delta * 4)
    heatmap[:, :, 1] = np.minimum(96, delta)
    diff_path = f"{prefix}-diff.png"
    Image.fromarray(heatmap).save(output / diff_path)
    source_render_path = f"{prefix}-compare.png"
    comparison = Image.new("RGB", (2 * width, height))
    comparison.paste(source, (0, 0))
    comparison.paste(rendered, (width, 0))
    comparison.save(output / source_render_path)
    return {
        "status": "compared", "overall": measure(delta),
        "outside_text": outside_text, "outside_text_and_charts": outside_text_and_charts,
        "outside_objects": outside_objects,
        "objects": regions, "outside_objects_tiles": tiles,
        "outside_objects_components": unassigned_changes, "review_flags": review_flags,
        "difference_image": diff_path, "side_by_side": source_render_path,
    }


def verify(manifest_path, pptx_path, outdir, renderer="auto", render_dir=None):
    manifest_path, pptx_path, output = (
        Path(path).resolve() for path in (manifest_path, pptx_path, outdir)
    )
    with manifest_path.open(encoding="utf-8-sig") as manifest_file:
        manifest = json.load(manifest_file)
    if manifest.get("version") != 1 or not manifest.get("slides"):
        raise ValueError("Expected a nonempty version 1 manifest.")
    presentation = Presentation(pptx_path)
    output.mkdir(parents=True, exist_ok=True)
    presentation_issues = []
    if len(presentation.slides) != len(manifest["slides"]):
        presentation_issues.append("PPT slide count differs from the manifest.")

    prepared = []
    render_specs = []
    for index, slide in enumerate(manifest["slides"]):
        source_path = (manifest_path.parent / slide["source"]).resolve()
        if Path(slide["source"]).is_absolute() or not source_path.is_relative_to(manifest_path.parent):
            raise ValueError(f"Slide {slide['id']!r} source is outside the manifest directory.")
        if slide.get("sha256") and hashlib.sha256(source_path.read_bytes()).hexdigest() != slide["sha256"]:
            raise ValueError(f"Slide {slide['id']!r} source checksum differs from the manifest.")
        with Image.open(source_path) as original:
            source = ImageOps.exif_transpose(original).convert("RGB")
        if source.size != (slide["width"], slide["height"]):
            raise ValueError(f"Slide {slide['id']!r} source dimensions differ from the manifest.")
        prefix = f"slide-{index + 1:03d}"
        inventory_name = f"{prefix}-inventory.png"
        annotated_inventory(source, slide["elements"]).save(output / inventory_name)
        rendered_path = output / f"{prefix}-render.png"
        prepared.append((slide, source, prefix, inventory_name, rendered_path))
        render_specs.append({
            "number": index + 1, "width": source.width, "height": source.height,
            "path": str(rendered_path),
        })

    with tempfile.TemporaryDirectory(prefix="slide-png-", dir=output) as rendered_files:
        for spec in render_specs:
            spec["path"] = str(Path(rendered_files, Path(spec["path"]).name))
        if render_dir is None:
            render_result = render_slides(pptx_path, output, render_specs, renderer)
        else:
            render_result = {"status": "exported", "backend": "supplied_pngs"}
        reports = []
        for index, (slide, source, prefix, inventory_name, rendered_path) in enumerate(prepared):
            report = check_slide(presentation, index, slide)
            report["inventory_image"] = inventory_name
            actual_render = Path(render_specs[index]["path"])
            if render_dir is not None:
                actual_render = Path(render_dir).resolve() / f"{prefix}.png"
            if actual_render.is_file():
                with Image.open(actual_render) as original:
                    rendered = original.convert("RGB")
                rendered.save(rendered_path)
                if rendered.size == source.size:
                    report["visual"] = compare_render(source, rendered, slide, prefix, output)
                else:
                    report["visual"] = {
                        "status": "size_mismatch", "source_size": source.size, "render_size": rendered.size,
                    }
                report["visual"]["render_image"] = rendered_path.name
            else:
                report["visual"] = {"status": "not_rendered"}
            reports.append(report)

    structural_error = presentation_issues or any(
        slide["issues"] or any(obj["issues"] for obj in slide["objects"]) for slide in reports
    )
    comparisons = sum(report["visual"]["status"] == "compared" for report in reports)
    if render_result["status"] == "exported":
        render_result["status"] = "complete" if comparisons == len(reports) else "incomplete"
    result = {
        "status": "draft_requires_review",
        "structural_status": "issues_found" if structural_error else "checked",
        "render": render_result, "slides_compared": comparisons,
        "pixel_delta_threshold": PIXEL_DELTA_THRESHOLD,
        "geometry_tolerance_px": STRUCTURE_TOLERANCE_PX,
        "issues": presentation_issues, "slides": reports,
    }
    (output / "report.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest")
    parser.add_argument("pptx")
    parser.add_argument("outdir")
    parser.add_argument("--renderer", choices=("auto", "powerpoint", "none"), default="auto")
    parser.add_argument("--render-dir", help="Compare supplied slide-001.png, slide-002.png, etc.")
    args = parser.parse_args()
    try:
        report = verify(args.manifest, args.pptx, args.outdir, args.renderer, args.render_dir)
    except (ValueError, OSError, KeyError, TypeError) as error:
        parser.exit(2, f"Verification could not finish: {error}\n")
    print(f"{Path(args.outdir).resolve() / 'report.json'}")
    print(f"Structure: {report['structural_status']}. Render: {report['render']['status']}.")
    if report["structural_status"] == "issues_found":
        return 1
    if (args.renderer == "powerpoint" or args.render_dir) and report["render"]["status"] != "complete":
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
