"""Build auditable, native PowerPoint charts with embedded Excel data."""

import math
import re
from sys import float_info

from pptx.chart.data import CategoryChartData, XyChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (
    XL_AXIS_CROSSES, XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION,
    XL_MARKER_STYLE, XL_TICK_LABEL_POSITION, XL_TICK_MARK,
)
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu


CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "line": XL_CHART_TYPE.LINE,
    "scatter": XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS,
}
DASHES = {
    "solid": MSO_LINE_DASH_STYLE.SOLID,
    "dash": MSO_LINE_DASH_STYLE.DASH,
    "dot": MSO_LINE_DASH_STYLE.ROUND_DOT,
}
MARKERS = {"none": XL_MARKER_STYLE.NONE, "circle": XL_MARKER_STYLE.CIRCLE}
PALETTE = ["#147D76", "#BE623A", "#465C8B", "#7A6057", "#657847"]
HEX = re.compile(r"#[0-9A-Fa-f]{6}\Z")
WORKBOOK_SPECIAL_TEXT = re.compile(r"^(?:=|\{=|https?://|ftp://|mailto:|file://|(?:in|ex)ternal:)", re.I)


def _plain(value, label, *, workbook_literal=False):
    if (not isinstance(value, str) or not value.strip() or len(value) > 32767
            or any(ord(char) < 32 or 0xD800 <= ord(char) <= 0xDFFF or char in "\ufffe\uffff" for char in value)):
        raise ValueError(f"{label}: expected nonempty plain text")
    if workbook_literal and WORKBOOK_SPECIAL_TEXT.match(value):
        raise ValueError(f"{label}: use a literal label; Excel formulas and links are unsupported")


def _numeric(value, label):
    if type(value) not in (int, float):
        raise ValueError(f"{label}: expected a finite Excel number")
    # Python can parse integers beyond Excel's floating-point range. Numeric
    # chart values also cannot rely on subnormals, which Excel may read as 0.
    if (abs(value) > float_info.max or not math.isfinite(value)
            or 0 < abs(value) < float_info.min or type(value) is int and abs(value) > 999_999_999_999_999):
        raise ValueError(f"{label}: outside Excel numeric range or exact integer precision")


def _numbers(values, label):
    if not isinstance(values, list) or not values:
        raise ValueError(f"{label}: expected a nonempty array of numbers")
    for idx, value in enumerate(values, 1):
        _numeric(value, f"{label}[{idx}]")


def _color(options, key, label):
    if key in options and (not isinstance(options[key], str) or not HEX.fullmatch(options[key])):
        raise ValueError(f"{label}/{key}: expected #RRGGBB")


def _flag(options, key, label):
    if key in options and type(options[key]) is not bool:
        raise ValueError(f"{label}/{key}: expected true or false")


def _axis(options, label, values, *, bars=False):
    if not isinstance(options, dict):
        raise ValueError(f"{label}: expected an axis object")
    unknown = options.keys() - {
        "min", "max", "major_unit", "title", "number_format",
        "show_major_gridlines", "visible", "show_tick_labels",
    }
    if unknown:
        raise ValueError(f"{label}: unsupported axis options: {', '.join(sorted(unknown))}")
    for key in ("min", "max", "major_unit"):
        if key in options:
            _numeric(options[key], f"{label}/{key}")
    if "min" in options and "max" in options and options["min"] >= options["max"]:
        raise ValueError(f"{label}: axis minimum must be less than maximum")
    if "major_unit" in options and options["major_unit"] <= 0:
        raise ValueError(f"{label}/major_unit: interval must be positive")
    if ("min" in options and "max" in options and "major_unit" in options
            and (options["max"] - options["min"]) / options["major_unit"] > 200):
        raise ValueError(f"{label}/major_unit: interval implies more than 200 major ticks")
    minimum = min(min(values), 0) if bars else min(values)
    maximum = max(max(values), 0) if bars else max(values)
    if options.get("min", minimum) > minimum or options.get("max", maximum) < maximum:
        detail = " (including the bar baseline at zero)" if bars else ""
        raise ValueError(f"{label}: axis bounds would hide data{detail}")
    for key in ("title", "number_format"):
        if key in options:
            _plain(options[key], f"{label}/{key}")
    for key in ("show_major_gridlines", "visible", "show_tick_labels"):
        _flag(options, key, label)


def validate_chart(element, label=None):
    """Require explicit labels, source provenance, aligned numeric series and honest bounds."""
    label = label or str(element.get("id", "chart"))
    chart_type = element.get("chart_type")
    if not isinstance(chart_type, str) or chart_type not in CHART_TYPES:
        raise ValueError(f"{label}/chart_type: expected column, bar, line, or scatter")
    xy = chart_type == "scatter"
    bars = chart_type in ("column", "bar")
    _plain(element.get("source_note"), f"{label}/source_note")
    for key in ("title", "font", "category_axis_title"):
        if key in element:
            _plain(element[key], f"{label}/{key}")
    for key in ("foreground", "background"):
        _color(element, key, label)
    for key in ("show_legend", "data_labels"):
        _flag(element, key, label)
    if "font_size_px" in element:
        size = element["font_size_px"]
        _numeric(size, f"{label}/font_size_px")
        if not 1 <= size <= 512:
            raise ValueError(f"{label}/font_size_px: expected 1–512 source pixels")

    categories = element.get("categories")
    if xy:
        if any(key in element for key in ("categories", "category_axis_title")):
            raise ValueError(f"{label}: scatter uses numeric x_values and optional x_axis, without categories")
    elif not isinstance(categories, list) or not 1 <= len(categories) <= 1_048_575:
        raise ValueError(f"{label}/categories: expected a nonempty label array")
    else:
        for idx, category in enumerate(categories, 1):
            _plain(category, f"{label}/categories[{idx}]", workbook_literal=True)
    if not xy and "x_axis" in element:
        raise ValueError(f"{label}/x_axis: numeric x_axis options are only supported for scatter")

    series = element.get("series")
    if not isinstance(series, list) or not 1 <= len(series) <= 255:
        raise ValueError(f"{label}/series: expected 1–255 series")
    names, x_values, y_values = set(), [], []
    for idx, row in enumerate(series, 1):
        series_label = f"{label}/series[{idx}]"
        if not isinstance(row, dict):
            raise ValueError(f"{series_label}: expected an object")
        _plain(row.get("name"), f"{series_label}/name", workbook_literal=True)
        if row["name"] in names:
            raise ValueError(f"{series_label}/name: duplicate series {row['name']!r}")
        names.add(row["name"])
        values = row.get("values")
        _numbers(values, f"{series_label}/values")
        y_values.extend(values)
        if xy:
            points = row.get("x_values")
            _numbers(points, f"{series_label}/x_values")
            if len(points) != len(values):
                raise ValueError(f"{series_label}: x_values and values must have the same length")
            x_values.extend(points)
        elif "x_values" in row or len(values) != len(categories):
            raise ValueError(f"{series_label}: categorical series must match categories and omit x_values")
        _color(row, "color", series_label)
        line_options = row.keys() & {"line_width_px", "dash", "marker"}
        if bars and line_options:
            raise ValueError(f"{series_label}: line-only options on a bar or column: {', '.join(sorted(line_options))}")
        if "dash" in row and row["dash"] not in tuple(DASHES):
            raise ValueError(f"{series_label}/dash: expected solid, dash, or dot")
        if "marker" in row and row["marker"] not in tuple(MARKERS):
            raise ValueError(f"{series_label}/marker: expected none or circle")
        if "line_width_px" in row:
            width = row["line_width_px"]
            _numeric(width, f"{series_label}/line_width_px")
            if not 0 <= width <= 100:
                raise ValueError(f"{series_label}/line_width_px: expected 0–100 source pixels")
        if not bars and (len(values) == 1 or row.get("line_width_px") == 0) and row.get("marker") != "circle":
            raise ValueError(f"{series_label}: a single or unconnected point needs a marker")
    if xy:
        if len(x_values) + 2 * len(series) - 1 > 1_048_576:
            raise ValueError(f"{label}: scatter data exceed the embedded Excel sheet")
        _axis(element.get("x_axis", {}), f"{label}/x_axis", x_values)
    _axis(element.get("value_axis", {}), f"{label}/value_axis", y_values, bars=bars)


def _rgb(value):
    return RGBColor.from_string(value[1:].upper())


def _font(font, face, size, foreground):
    font.name = face
    # DrawingML can use a different East Asian typeface from the Latin one.
    rpr = font._rPr
    east_asian = rpr.find(qn("a:ea"))
    if east_asian is None:
        east_asian = OxmlElement("a:ea")
        rpr.insert_element_before(east_asian, "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:extLst")
    east_asian.set("typeface", face)
    font.size = size
    font.color.rgb = foreground


def _frame_text(frame, label, face, size, foreground):
    frame.text = label
    for paragraph in frame.paragraphs:
        for run in paragraph.runs:
            _font(run.font, face, size, foreground)


def _background(chart, color):
    # python-pptx exposes chart font/axes but currently not chart-area fill.
    chart_space = chart.element
    shape_properties = OxmlElement("c:spPr")
    fill = OxmlElement("a:solidFill") if color else OxmlElement("a:noFill")
    if color:
        node = OxmlElement("a:srgbClr")
        node.set("val", color[1:].upper())
        fill.append(node)
    shape_properties.append(fill)
    line = OxmlElement("a:ln")
    line.append(OxmlElement("a:noFill"))
    shape_properties.append(line)
    chart_space.insert_element_before(shape_properties, "c:txPr", "c:externalData", "c:printSettings",
                                      "c:userShapes", "c:extLst")


def _style_axis(axis, settings, face, font_size, foreground, *, numeric=False, grid=False):
    _font(axis.tick_labels.font, face, font_size, foreground)
    axis.major_tick_mark = XL_TICK_MARK.NONE
    axis.minor_tick_mark = XL_TICK_MARK.NONE
    axis.format.line.color.rgb = foreground
    if "title" in settings:
        axis.has_title = True
        _frame_text(axis.axis_title.text_frame, settings["title"], face, font_size, foreground)
    if numeric:
        if "visible" in settings:
            axis.visible = settings["visible"]
        if settings.get("show_tick_labels") is False:
            axis.tick_label_position = XL_TICK_LABEL_POSITION.NONE
        axis.has_major_gridlines = settings.get("show_major_gridlines", grid)
        if axis.has_major_gridlines:
            axis.major_gridlines.format.line.color.rgb = RGBColor.from_string("D7DFE0")
        for field, target in (("min", "minimum_scale"), ("max", "maximum_scale"), ("major_unit", "major_unit")):
            if field in settings:
                setattr(axis, target, settings[field])
        if "number_format" in settings:
            axis.tick_labels.number_format = settings["number_format"]
            axis.tick_labels.number_format_is_linked = False


def _chart_data(element, *, xy):
    formatted_values = element.get("value_axis", {}).get("number_format")
    if xy:
        data = XyChartData(number_format=element.get("x_axis", {}).get("number_format", "General"))
        for row in element["series"]:
            series = data.add_series(row["name"], number_format=formatted_values)
            for x, y in zip(row["x_values"], row["values"]):
                series.add_data_point(x, y)
        return data
    data = CategoryChartData()
    data.categories = element["categories"]
    for row in element["series"]:
        data.add_series(row["name"], row["values"], number_format=formatted_values)
    return data


def add_chart(slide, element, box, scale):
    """Place an editable chart with native axes, series, and embedded workbook."""
    chart_type = element["chart_type"]
    xy = chart_type == "scatter"
    bars = chart_type in ("column", "bar")
    frame = slide.shapes.add_chart(CHART_TYPES[chart_type], *box, _chart_data(element, xy=xy))
    frame.name = element["id"]
    frame._element.nvGraphicFramePr.cNvPr.set("descr", element["source_note"])
    chart = frame.chart
    face = element.get("font", "Microsoft YaHei")
    size = Emu(max(12700, round(element.get("font_size_px", 16) * scale)))
    foreground = _rgb(element.get("foreground", "#334155"))
    _font(chart.font, face, size, foreground)
    _background(chart, element.get("background"))
    chart.has_title = "title" in element
    if chart.has_title:
        _frame_text(chart.chart_title.text_frame, element["title"], face, size, foreground)
    chart.has_legend = element.get("show_legend", len(element["series"]) > 1)
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        _font(chart.legend.font, face, size, foreground)
    value_options = element.get("value_axis", {})
    if xy:
        _style_axis(chart.category_axis, element.get("x_axis", {}), face, size, foreground, numeric=True)
        chart.category_axis.crosses = XL_AXIS_CROSSES.MINIMUM
        chart.value_axis.crosses = XL_AXIS_CROSSES.MINIMUM
    else:
        title = element.get("category_axis_title")
        _style_axis(chart.category_axis, {"title": title} if title is not None else {}, face, size, foreground)
        if chart_type == "bar":
            chart.category_axis.reverse_order = True
            chart.category_axis.crosses = XL_AXIS_CROSSES.MAXIMUM
    _style_axis(chart.value_axis, value_options, face, size, foreground, numeric=True, grid=True)
    chart.plots[0].vary_by_categories = False
    for idx, (series, row) in enumerate(zip(chart.series, element["series"])):
        color = _rgb(row.get("color", PALETTE[idx % len(PALETTE)]))
        if bars:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = color
            series.format.line.fill.background()
            continue
        line = series.format.line
        width = row.get("line_width_px", 3)
        if width == 0:
            line.fill.background()
        else:
            line.color.rgb = color
            line.width = Emu(max(1, round(width * scale)))
            line.dash_style = DASHES[row.get("dash", "solid")]
        marker = series.marker
        marker.style = MARKERS[row.get("marker", "none")]
        if marker.style == XL_MARKER_STYLE.CIRCLE:
            marker.size = max(2, min(72, round(size.pt * 0.6)))
            marker.format.fill.solid()
            marker.format.fill.fore_color.rgb = color
            marker.format.line.fill.background()
    if element.get("data_labels", False):
        plot = chart.plots[0]
        plot.has_data_labels = True
        labels = plot.data_labels
        labels.show_value = True
        labels.show_category_name = labels.show_series_name = False
        labels.position = XL_LABEL_POSITION.OUTSIDE_END if bars else XL_LABEL_POSITION.ABOVE
        _font(labels.font, face, size, foreground)
        if "number_format" in value_options:
            labels.number_format = value_options["number_format"]
            labels.number_format_is_linked = False
    return frame
