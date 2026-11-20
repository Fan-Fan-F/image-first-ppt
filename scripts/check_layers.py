#!/usr/bin/env python3
"""Move individual PowerPoint graphics to neutral review slides."""

import argparse
from copy import deepcopy
from io import BytesIO
import json
from pathlib import Path
from tempfile import TemporaryDirectory

import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Inches

from raster import _element_mask, prepare
from rebuild import load_scene
from verify import DRAWING_TYPES, NATIVE_GEOMETRIES, STRUCTURE_TOLERANCE_PX, render_slides, shape_bounds


R_NS = qn("r:embed").split("}")[0] + "}"
REVIEW_COLOR = "#ECEFF1"
GRAPHICS = {"image", "shape"}
TILE = (332, 190)
ROWS_PER_SHEET = 6


def intersect(a, b):
    x, y = max(a[0], b[0]), max(a[1], b[1])
    right = min(a[0] + a[2], b[0] + b[2])
    bottom = min(a[1] + a[3], b[1] + b[3])
    return (x, y, right, bottom) if x < right and y < bottom else None


def index_build_report(report_path):
    if report_path is None:
        return {}
    raw = json.loads(report_path.read_text(encoding="utf-8-sig"))
    return {slide["id"]: slide.get("raster", {}).get("visuals", {}) for slide in raw["slides"]}


def select_objects(scene, requested):
    known = {}
    overlaps = set()
    for slide_index, slide in enumerate(scene["slides"]):
        graphics = [item for item in slide["elements"] if item["kind"] in GRAPHICS]
        for element_index, item in enumerate(graphics):
            key = f"{slide['id']}/{item['id']}"
            covers = [other for other in graphics[element_index + 1:] if intersect(item["box"], other["box"])]
            known[key] = (slide_index, slide, item, covers)
            if covers or "completion" in item:
                overlaps.add(key)
                overlaps.update(f"{slide['id']}/{cover['id']}" for cover in covers)
    if not known:
        raise ValueError("the manifest has no image or native shape to inspect")
    missing = sorted(set(requested) - known.keys())
    if missing:
        raise ValueError(f"unknown graphic IDs: {', '.join(missing)}")
    keys = list(dict.fromkeys(requested)) if requested else [key for key in known if key in overlaps] or list(known)
    return [known[key] for key in keys]


def covered_pixels(item, covers, canvas_size, root):
    size = item["box"][3], item["box"][2]
    upper = np.zeros(size, dtype=bool)
    local_region, declared = _element_mask(item, canvas_size, root)
    rows, columns = local_region
    for cover in covers:
        (above_rows, above_columns), above = _element_mask(cover, canvas_size, root)
        left = max(columns.start, above_columns.start)
        top = max(rows.start, above_rows.start)
        right = min(columns.stop, above_columns.stop)
        bottom = min(rows.stop, above_rows.stop)
        if left < right and top < bottom:
            upper[top - rows.start:bottom - rows.start, left - columns.start:right - columns.start] |= (
                above[top - above_rows.start:bottom - above_rows.start,
                      left - above_columns.start:right - above_columns.start]
            )
    return upper, upper & declared


def bitmap_stats(bitmap, item, upper, declared):
    alpha = np.asarray(bitmap.getchannel("A"))
    present = alpha > 0
    strong = alpha >= 128
    count, _, stats, _ = cv2.connectedComponentsWithStats(strong.astype(np.uint8), connectivity=8)
    areas = sorted((int(row[cv2.CC_STAT_AREA]) for row in stats[1:count]), reverse=True)
    _, voids, void_stats, _ = cv2.connectedComponentsWithStats((~strong).astype(np.uint8), connectivity=4)
    exterior = set(np.concatenate((voids[0], voids[-1], voids[:, 0], voids[:, -1])))
    holes = [int(row[cv2.CC_STAT_AREA]) for label, row in enumerate(void_stats) if label and label not in exterior]
    expected_size = tuple(item["box"][2:])
    measures = {
        "embedded_size_px": list(bitmap.size),
        "alpha_bounds_px": bitmap.getchannel("A").getbbox(),
        "alpha_nonzero_pixels": int(present.sum()),
        "alpha_at_least_half_pixels": int(strong.sum()),
        "connected_components_at_half_alpha": len(areas),
        "largest_component_areas_px": areas[:10],
        "enclosed_transparent_regions_at_half_alpha": len(holes),
        "largest_enclosed_region_areas_px": sorted(holes, reverse=True)[:10],
        "declared_support_behind_upper_layers_px": int(declared.sum()),
    }
    if bitmap.size == expected_size:
        measures.update({
            "embedded_support_behind_upper_layers_px": int((present & upper).sum()),
            "declared_support_missing_from_embedded_alpha_px": int((declared & ~present).sum()),
        })
    return measures


def compare_bitmap(actual, expected):
    if expected is None or expected.size != actual.size:
        return {"status": "size_or_expected_asset_mismatch"}
    exported, planned = np.asarray(actual), np.asarray(expected)
    actual_alpha, expected_alpha = exported[:, :, 3], planned[:, :, 3]
    both_present = (actual_alpha > 0) & (expected_alpha > 0)
    changed_color = np.any(exported[:, :, :3] != planned[:, :, :3], axis=2) & both_present
    different_alpha = int((actual_alpha != expected_alpha).sum())
    return {
        "status": "matches" if not different_alpha and not changed_color.any() else "different",
        "alpha_difference_px": different_alpha,
        "missing_alpha_support_px": int(((expected_alpha > 0) & (actual_alpha == 0)).sum()),
        "extra_alpha_support_px": int(((expected_alpha == 0) & (actual_alpha > 0)).sum()),
        "color_difference_inside_common_alpha_px": int(changed_color.sum()),
    }


def neutral_slide(deck, layout):
    page = deck.slides.add_slide(layout)
    for placeholder in list(page.placeholders):
        placeholder.element.getparent().remove(placeholder.element)
    page._element.set("showMasterSp", "0")
    page.background.fill.solid()
    page.background.fill.fore_color.rgb = RGBColor.from_string(REVIEW_COLOR[1:])
    return page


def clone_graphic(source, destination, expected_kind):
    shape_type = source.shape_type
    if expected_kind == "image" and shape_type != MSO_SHAPE_TYPE.PICTURE:
        raise ValueError(f"{source.name}: expected a PowerPoint picture")
    if expected_kind == "shape" and shape_type not in DRAWING_TYPES:
        raise ValueError(f"{source.name}: expected a native PowerPoint drawing")

    element = deepcopy(source.element)
    embedded = element.find(".//" + qn("a:blip")) if expected_kind == "image" else None
    for node in element.iter():
        for attribute in node.attrib:
            if attribute.startswith(R_NS) and not (node is embedded and attribute == qn("r:embed")):
                raise ValueError(f"{source.name}: drawing requires an unsupported linked resource")
    if expected_kind == "image":
        image_part, image_rid = destination.part.get_or_add_image_part(BytesIO(source.image.blob))
        if image_part.blob != source.image.blob or embedded is None:
            raise ValueError(f"{source.name}: cannot reproduce the embedded picture")
        embedded.set(qn("r:embed"), image_rid)

    identity = element.find(".//" + qn("p:cNvPr"))
    if identity is None:
        raise ValueError(f"{source.name}: drawing has no identity element")
    identity.set("id", str(destination.shapes._next_shape_id))
    destination.shapes._spTree.insert_element_before(element, "p:extLst")
    return destination.shapes[-1]


def inspect_object(deck, original_pages, selection, source_paths, build_index, root, evidence):
    slide_index, slide, item, covers = selection
    source_page = original_pages[slide_index]
    matching = [shape for shape in source_page.shapes if shape.name == item["id"]]
    identifier = f"{slide['id']}/{item['id']}"
    if len(matching) != 1:
        raise ValueError(f"{identifier}: expected exactly one graphic in the input deck")
    shape = matching[0]
    data = {
        "id": identifier,
        "kind": item["kind"],
        "expected_box_px": item["box"],
        "box_overlap_candidates_above": [cover["id"] for cover in covers],
        "issues": [],
    }
    expected_assets, planned_visuals = evidence
    expected_visual = planned_visuals.get(item["id"], {})
    recorded = build_index.get(slide["id"], {}).get(item["id"])
    if recorded is not None:
        data["build_report_evidence"] = recorded
    completion = expected_visual.get("completion") or (recorded or {}).get("completion")
    if completion:
        data["completion"] = completion
        if completion["status"] in {"incomplete_overlap", "possible_overlap"}:
            data["issues"].append("The manifest has not resolved possible hidden parts. Inspect on an isolated slide.")

    page = neutral_slide(deck, source_page.slide_layout)
    isolated = clone_graphic(shape, page, item["kind"])
    isolated.left = (deck.slide_width - isolated.width) // 2
    isolated.top = (deck.slide_height - isolated.height) // 2
    if (isolated.width, isolated.height, isolated.rotation) != (shape.width, shape.height, shape.rotation):
        raise ValueError(f"{identifier}: copied drawing differs in size or rotation")
    data["moves_as_one_object"] = len(page.shapes) == 1
    data["isolation_translation_emu"] = [isolated.left - shape.left, isolated.top - shape.top]
    data["review_box_emu"] = [isolated.left, isolated.top, isolated.width, isolated.height]
    canvas = (slide["width"], slide["height"])
    with Image.open(source_paths[slide_index]) as image:
        x, y, w, h = item["box"]
        source_crop = image.crop((x, y, x + w, y + h)).convert("RGB")

    image_preview = None
    if item["kind"] == "image":
        with Image.open(BytesIO(isolated.image.blob)) as embedded:
            bitmap = embedded.convert("RGBA")
        if isolated.image.blob != shape.image.blob:
            raise ValueError(f"{identifier}: image changed when copied")
        data["embedded_image_bytes_preserved"] = True
        upper, declared = covered_pixels(item, covers, canvas, root)
        data["alpha"] = alpha = bitmap_stats(bitmap, item, upper, declared)
        if alpha["alpha_nonzero_pixels"] == 0:
            data["issues"].append("Embedded object is completely transparent.")
        if bitmap.size != tuple(item["box"][2:]) or any(
                getattr(shape, f"crop_{side}", 0) for side in ("left", "top", "right", "bottom")):
            data["issues"].append("Embedded pixels do not map 1:1 to the full declared object box. Check crop and size.")
        if alpha.get("declared_support_missing_from_embedded_alpha_px") and (completion or {}).get("status") != "completed":
            data["issues"].append("Embedded alpha has gaps within declared support behind an upper object.")
        comparison = compare_bitmap(bitmap, expected_assets.get(item["id"]))
        data["manifest_asset_comparison"] = comparison
        if comparison["status"] != "matches":
            data["issues"].append("Embedded picture differs from the manifest's expected whole-object asset.")
        image_preview = ImageOps.contain(bitmap, TILE, method=Image.Resampling.LANCZOS)
    else:
        data["native_type"] = shape.shape_type.name
        actual_geometry = shape.auto_shape_type if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE else None
        data["native_geometry"] = actual_geometry.name if actual_geometry is not None else None
        if actual_geometry != NATIVE_GEOMETRIES.get(item["shape"]):
            data["issues"].append("Native geometry differs from the declared complete shape.")
    return data, ImageOps.contain(source_crop, TILE, method=Image.Resampling.LANCZOS), image_preview, shape


def checkerboard(bitmap):
    matte = Image.new("RGBA", bitmap.size, "#E1E6E9")
    draw = ImageDraw.Draw(matte)
    step = max(4, round(max(bitmap.size) / 18))
    for y in range(0, bitmap.height, step):
        for x in range((y // step % 2) * step, bitmap.width, 2 * step):
            draw.rectangle((x, y, x + step - 1, y + step - 1), fill="#FFF")
    matte.alpha_composite(bitmap)
    return matte.convert("RGB")


def contact_sheets(rows, out_dir, slide_size, rendered_dir=None):
    sheets = []
    width, height = TILE
    margin, header, row_header = 16, 29, 22
    for top in range(0, len(rows), ROWS_PER_SHEET):
        subset = rows[top:top + ROWS_PER_SHEET]
        contact = Image.new("RGB", (3 * margin + 2 * width, header + len(subset) * (height + row_header + margin)), "#FFF")
        draw = ImageDraw.Draw(contact)
        draw.text((margin, 8), "Source crop", fill="#22303C")
        draw.text((2 * margin + width, 8), "PowerPoint export" if rendered_dir else "Embedded object", fill="#22303C")
        for col, (data, source_crop, bitmap) in enumerate(subset):
            y = header + col * (height + row_header + margin)
            draw.text((margin, y), f"Page {data['review_slide']}   {data['id']}", fill="#22303C")
            draw.rectangle((margin, y + row_header, margin + width, y + row_header + height), fill=REVIEW_COLOR)
            draw.rectangle((2 * margin + width, y + row_header, 2 * margin + 2 * width, y + row_header + height),
                           fill=REVIEW_COLOR)
            if rendered_dir:
                with Image.open(rendered_dir / f"slide-{data['review_slide']:03}.png") as rendered_image:
                    x, y, w, h = data["review_box_emu"]
                    horizontal = rendered_image.width / slide_size[0]
                    vertical = rendered_image.height / slide_size[1]
                    margin_px = max(4, round(min(w * horizontal, h * vertical) * 0.05))
                    crop = (max(0, round(x * horizontal) - margin_px),
                            max(0, round(y * vertical) - margin_px),
                            min(rendered_image.width, round((x + w) * horizontal) + margin_px),
                            min(rendered_image.height, round((y + h) * vertical) + margin_px))
                    preview = rendered_image.crop(crop).convert("RGB")
            else:
                preview = checkerboard(bitmap) if bitmap else None
            for index, image in enumerate((source_crop, preview)):
                x = margin + index * (margin + width)
                if image:
                    fit = ImageOps.contain(image, TILE, method=Image.Resampling.LANCZOS)
                    contact.paste(fit, (x + (width - fit.width) // 2, y + row_header + (height - fit.height) // 2))
                else:
                    draw.text((x + 14, y + row_header + 14), "Open isolated.pptx to view this shape.", fill="#22303C")
        name = f"contact-{top // ROWS_PER_SHEET + 1:03}.png"
        contact.save(out_dir / name)
        sheets.append(name)
    return sheets


def export_review(deck_path, out_dir, count, slide_size, renderer):
    if renderer == "none":
        return render_slides(deck_path, out_dir, [], renderer), None
    width, height = slide_size
    export_scale = min(1280 / width, 800 / height)
    export_size = (round(width * export_scale), round(height * export_scale))
    with TemporaryDirectory(prefix="layer-review-render-", dir=out_dir) as temporary:
        temporary = Path(temporary)
        render_plan = [
            {"number": n, "path": str(temporary / f"slide-{n:03}.png"),
             "width": export_size[0], "height": export_size[1]}
            for n in range(1, count + 1)
        ]
        render = render_slides(deck_path, out_dir, render_plan, renderer)
        if render.get("status") != "exported":
            return render, None
        frames = [Path(frame["path"]) for frame in render_plan]
        if not all(frame.is_file() for frame in frames):
            return {"status": "failed", "reason": "PowerPoint did not produce all review frames."}, None
        published = out_dir / "render"
        published.mkdir(exist_ok=True)
        for frame in frames:
            frame.replace(published / frame.name)
        return render, published


def protect_inputs(scene, manifest, source_paths, pptx_path, report_path, targets):
    inputs = {manifest.resolve(), pptx_path.resolve(), *(path.resolve() for path in source_paths)}
    if report_path:
        inputs.add(report_path.resolve())
    for slide in scene["slides"]:
        for element in slide["elements"]:
            completion = element.get("completion", {})
            for file in (element.get("mask"), element.get("erase_mask"), element.get("erase_patch"),
                         element.get("underlay_patch"), completion.get("file"), completion.get("mask")):
                if file:
                    inputs.add((manifest.parent / file).resolve())
    if any(target.resolve() in inputs for target in targets):
        raise ValueError("choose a review output directory that does not overlap the input files")


def build_review(manifest, pptx_path, out_dir, requested, renderer, report_path=None):
    manifest = manifest.resolve()
    scene, source_paths = load_scene(manifest)
    selected = select_objects(scene, requested)
    if report_path is None and pptx_path.with_suffix(".build.json").is_file():
        report_path = pptx_path.with_suffix(".build.json")
    build_index = index_build_report(report_path)
    deck = Presentation(pptx_path)
    original_pages = tuple(deck.slides)
    if len(original_pages) != len(scene["slides"]):
        raise ValueError("slide count differs between the manifest and the input PowerPoint")

    original_width, original_height = deck.slide_width, deck.slide_height
    scale = min(1.2, Inches(56) / max(original_width, original_height))
    deck.slide_width, deck.slide_height = round(original_width * scale), round(original_height * scale)
    rows = []
    prepared_slide = None
    evidence = ({}, {})
    for page_number, selection in enumerate(selected, 1):
        slide_index, slide, item, _ = selection
        if slide_index != prepared_slide and item["kind"] == "image":
            with Image.open(source_paths[slide_index]) as original:
                _, expected_assets, stats, _ = prepare(slide, original.convert("RGB"), manifest.parent)
            evidence = ({element["id"]: bitmap for element, bitmap in expected_assets}, stats["visuals"])
            prepared_slide = slide_index
        page_evidence = evidence if prepared_slide == slide_index else ({}, {})
        data, source_crop, bitmap, shape = inspect_object(
            deck, original_pages, selection, source_paths, build_index, manifest.parent, page_evidence
        )
        data["review_slide"] = page_number
        actual = shape_bounds(shape, original_width, original_height, slide["width"], slide["height"])
        data["input_box_px"] = [round(value, 3) for value in actual] if actual else None
        if not actual or max(abs(a - b) for a, b in zip(actual, data["expected_box_px"])) > STRUCTURE_TOLERANCE_PX:
            data["issues"].append("Input object position or size differs from the manifest.")
        rows.append((data, source_crop, bitmap))

    slide_ids = deck.slides._sldIdLst
    for slide_id in list(slide_ids)[:len(original_pages)]:
        deck.part.drop_rel(slide_id.rId)
        slide_ids.remove(slide_id)
    out_dir = out_dir.resolve()
    isolated_deck = out_dir / "isolated.pptx"
    report_file = out_dir / "report.json"
    targets = [isolated_deck, report_file]
    targets.extend(out_dir / f"contact-{i:03}.png" for i in range(1, (len(rows) - 1) // ROWS_PER_SHEET + 2))
    if renderer != "none":
        targets.extend(out_dir / "render" / f"slide-{i:03}.png" for i in range(1, len(rows) + 1))
    protect_inputs(scene, manifest, source_paths, pptx_path, report_path, targets)
    out_dir.mkdir(parents=True, exist_ok=True)
    deck.save(isolated_deck)
    slide_size = (deck.slide_width, deck.slide_height)
    render, rendered_dir = export_review(isolated_deck, out_dir, len(rows), slide_size, renderer)
    sheets = contact_sheets(rows, out_dir, slide_size, rendered_dir)
    if rendered_dir:
        for data, _, _ in rows:
            data["rendered_png"] = f"render/slide-{data['review_slide']:03}.png"
    report = {
        "status": "manual_layer_review_required",
        "isolated_deck": isolated_deck.name,
        "contact_sheets": sheets,
        "renderer": render,
        "preview_kind": "powerpoint" if rendered_dir else "embedded_preview",
        "interpretation": (
            "Transparency, disconnected parts and holes may be intentional. Declared overlap metrics use source masks "
            "and cannot establish the actual original appearance of hidden pixels. Inspect the isolated objects."
        ),
        "objects": [data for data, _, _ in rows],
    }
    report_file.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return report_file, rendered_dir is not None, any(data["issues"] for data, _, _ in rows)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path)
    parser.add_argument("pptx", type=Path)
    parser.add_argument("out_dir", type=Path)
    parser.add_argument("--object", action="append", default=[], help="which slide/id to isolate; repeat for each")
    parser.add_argument("--renderer", choices=("none", "auto", "powerpoint"), default="none")
    parser.add_argument("--build-report", type=Path, help="override the sibling .build.json evidence")
    parser.add_argument("--require-complete", action="store_true",
                        help="fail if selected objects have unresolved support, geometry, or image mismatches")
    args = parser.parse_args()
    try:
        path, rendered, has_issues = build_review(
            args.manifest, args.pptx, args.out_dir, args.object, args.renderer, args.build_report
        )
    except (OSError, ValueError, KeyError, TypeError) as error:
        parser.exit(2, f"Layer review stopped: {error}\n")
    print(path)
    if (args.renderer == "powerpoint" and not rendered) or (args.require_complete and has_issues):
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
