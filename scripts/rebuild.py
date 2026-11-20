"""Freeze source slides and assemble an editable PowerPoint from a scene manifest."""

import argparse
import hashlib
import io
import json
import math
import os
import re
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

from raster import prepare
from typography import add_text


SHAPES = {
    "rect": MSO_SHAPE.RECTANGLE,
    "round_rect": MSO_SHAPE.ROUNDED_RECTANGLE,
    "ellipse": MSO_SHAPE.OVAL,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
}
HEX = re.compile(r"#[0-9A-Fa-f]{6}\Z")
IDENTIFIER = re.compile(r"[a-zA-Z0-9][a-zA-Z0-9_-]{0,63}\Z")


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def source_path(root, name):
    if not isinstance(name, str) or not name or Path(name).is_absolute():
        raise ValueError("source must be a relative file path")
    path = (root / name).resolve()
    if not path.is_relative_to(root.resolve()) or not path.is_file():
        raise ValueError(f"source is outside the manifest folder or missing: {name}")
    return path


def check_box(box, width, height, label):
    if not isinstance(box, list) or len(box) != 4 or any(type(v) is not int for v in box):
        raise ValueError(f"{label}: expected four integer pixel coordinates")
    x, y, w, h = box
    if x < 0 or y < 0 or w <= 0 or h <= 0 or x + w > width or y + h > height:
        raise ValueError(f"{label}: outside source canvas or empty: {box}")


def check_polygon(points, box, label):
    if not isinstance(points, list) or len(points) < 3:
        raise ValueError(f"{label}: expected at least three polygon points")
    x, y, w, h = box
    for point in points:
        if (not isinstance(point, list) or len(point) != 2
                or any(type(v) is not int for v in point)
                or not (x <= point[0] < x + w and y <= point[1] < y + h)):
            raise ValueError(f"{label}: polygon vertex outside its pixel box: {point}")


def check_id(value, known, label):
    if not isinstance(value, str) or not IDENTIFIER.fullmatch(value) or value in known:
        raise ValueError(f"{label}: invalid or duplicate id: {value!r}")
    known.add(value)


def number(value, low, high, label):
    if type(value) not in (float, int) or not math.isfinite(value) or not low <= value <= high:
        raise ValueError(f"{label}: expected a number between {low} and {high}")


def choice(obj, key, options, label):
    if key in obj and obj[key] not in options:
        raise ValueError(f"{label}: unknown {key}: {obj[key]}")


def load_scene(path):
    scene = json.loads(path.read_text(encoding="utf-8-sig"))
    if scene.get("version") != 1 or not isinstance(scene.get("slides"), list) or not scene["slides"]:
        raise ValueError("expected a version 1 manifest with one or more slides")
    number(scene.get("slide_width_in", 13.33333333), 2, 56, "slide_width_in")
    known = set()
    sources = []
    for slide in scene["slides"]:
        slide_id = slide.get("id")
        check_id(slide_id, known, "slide")
        file = source_path(path.parent, slide.get("source"))
        expected = slide.get("sha256")
        if expected and expected != digest(file):
            raise ValueError(f"{slide_id}: frozen source hash changed: {file.name}")
        with Image.open(file) as raw:
            if raw.format != "PNG" or raw.mode != "RGB" or raw.getexif().get(274, 1) != 1:
                raise ValueError(f"{slide_id}: freeze the source as an oriented RGB PNG with init")
            width, height = raw.size
        if (width, height) != (slide.get("width"), slide.get("height")):
            raise ValueError(f"{slide_id}: manifest dimensions do not match frozen source")
        elements = slide.get("elements")
        if not isinstance(elements, list) or not elements:
            raise ValueError(f"{slide_id}: element inventory is empty")
        item_ids = set()
        for item in elements:
            item_id = item.get("id")
            check_id(item_id, item_ids, slide_id)
            label = f"{slide_id}/{item_id}"
            kind = item.get("kind")
            if kind not in ("text", "image", "shape"):
                raise ValueError(f"{label}: unknown kind: {kind!r}")
            box = item.get("box")
            check_box(box, width, height, label)
            for key in ("color", "erase_fill", "underlay_fill", "fill", "stroke"):
                if key in item and (not isinstance(item[key], str) or not HEX.fullmatch(item[key])):
                    raise ValueError(f"{label}: {key} must be #RRGGBB")
            if kind == "text":
                if not isinstance(item.get("text"), str) or not item["text"].strip():
                    raise ValueError(f"{label}: editable text is empty")
                if any(ord(c) < 32 and c not in "\n\t" for c in item["text"]):
                    raise ValueError(f"{label}: use LF line breaks and remove control characters")
                erase = item.get("erase", box)
                check_box(erase, width, height, f"{label}/erase")
                if "erase_polygon" in item:
                    check_polygon(item["erase_polygon"], erase, f"{label}/erase_polygon")
                if "erase_mask" in item:
                    source_path(path.parent, item["erase_mask"])
                    if "erase_polygon" in item:
                        raise ValueError(f"{label}: choose erase_mask or erase_polygon")
                if "erase_patch" in item:
                    source_path(path.parent, item["erase_patch"])
                    if "erase_fill" in item or not ("erase_mask" in item or "erase_polygon" in item):
                        raise ValueError(f"{label}: erase_patch requires a tight erase_mask or erase_polygon, without erase_fill")
                choice(item, "align", ("left", "center", "right"), label)
                choice(item, "valign", ("top", "middle", "bottom"), label)
                choice(item, "erase_mode", ("auto", "flat", "inpaint"), label)
                for key in ("font", "font_ea"):
                    if key in item and (not isinstance(item[key], str) or not item[key].strip()):
                        raise ValueError(f"{label}: {key} must be a font family name")
                for key in ("bold", "italic"):
                    if key in item and type(item[key]) is not bool:
                        raise ValueError(f"{label}: {key} must be true or false")
                for key, limits in {"font_px": (1, 4096), "line_height_px": (1, 4096),
                                    "rotation": (-359.999, 359.999)}.items():
                    if key in item:
                        number(item[key], *limits, f"{label}/{key}")
            else:
                if "polygon" in item:
                    check_polygon(item["polygon"], box, f"{label}/polygon")
                if "mask" in item:
                    source_path(path.parent, item["mask"])
                if "underlay_patch" in item:
                    source_path(path.parent, item["underlay_patch"])
                    if "underlay_fill" in item:
                        raise ValueError(f"{label}: choose underlay_patch or underlay_fill")
                if kind == "shape":
                    if item.get("shape") not in SHAPES or "fill" not in item:
                        raise ValueError(f"{label}: native shape requires supported shape and fill")
                    if "stroke_px" in item:
                        number(item["stroke_px"], 0.05, 2048, f"{label}/stroke_px")
                if item.get("rotation", 0):
                    raise ValueError(f"{label}: use a source-pixel image mask for rotated artwork")
        sources.append(file)
    return scene, sources


def picture(slide, image, box, name):
    with io.BytesIO() as stream:
        image.save(stream, format="PNG")
        stream.seek(0)
        shape = slide.shapes.add_picture(stream, *box)
    shape.name = name


def native_shape(slide, item, box, scale):
    shape = slide.shapes.add_shape(SHAPES[item["shape"]], *box)
    shape.name = item["id"]
    style = shape._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}style")
    if style is not None:
        shape._element.remove(style)
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(item["fill"][1:].upper())
    if "stroke" in item:
        shape.line.color.rgb = RGBColor.from_string(item["stroke"][1:].upper())
        shape.line.width = max(1, round(item.get("stroke_px", 1) * scale))
    else:
        shape.line.fill.background()


def write_json(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def initialize(job_dir, inputs, slide_width):
    number(slide_width, 2, 56, "slide width")
    job_dir.mkdir(parents=True, exist_ok=True)
    manifest = job_dir / "manifest.json"
    if manifest.exists():
        raise ValueError(f"manifest already exists: {manifest}")
    slides = []
    for i, path in enumerate(inputs, 1):
        slide_id = f"s{i:03}"
        target = job_dir / f"{slide_id}.png"
        if target.exists():
            raise ValueError(f"refusing to overwrite frozen source: {target}")
        with Image.open(path) as raw:
            source = ImageOps.exif_transpose(raw).convert("RGBA")
        plate = Image.new("RGBA", source.size, "white")
        plate.alpha_composite(source)
        plate.convert("RGB").save(target, format="PNG")
        slides.append({"id": slide_id, "source": target.name, "width": source.width,
                       "height": source.height, "sha256": digest(target), "elements": []})
    write_json(manifest, {"version": 1, "slide_width_in": slide_width, "slides": slides})
    return manifest


def build(manifest, output, layers_dir=None):
    manifest = manifest.resolve()
    scene, sources = load_scene(manifest)
    first = scene["slides"][0]
    width, height = first["width"], first["height"]
    deck = Presentation()
    deck.slide_width = Inches(scene.get("slide_width_in", 13.33333333))
    deck.slide_height = round(deck.slide_width * height / width)
    if not Inches(1) <= deck.slide_height <= Inches(56):
        raise ValueError("slide height calculated from the source aspect ratio is outside 1–56 inches")
    results = []
    for data, path in zip(scene["slides"], sources):
        with Image.open(path) as raw:
            source = raw.convert("RGB")
        if source.width * height != source.height * width:
            raise ValueError(f"{data['id']}: mixed aspect ratios in a single deck. Pad/crop deliberately before init.")
        scale = deck.slide_width / source.width
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        background, rasters, raster_stats, textless = prepare(data, source, manifest.parent)
        raster_by_id = {item["id"]: image for item, image in rasters}
        if layers_dir:
            layer_folder = layers_dir / data["id"]
            layer_folder.mkdir(parents=True, exist_ok=True)
            background.save(layer_folder / "__background__.png")
            textless.save(layer_folder / "__textless_source__.png")
            for item_id, raster in raster_by_id.items():
                raster.save(layer_folder / f"{item_id}.png")
        picture(slide, background, (0, 0, deck.slide_width, deck.slide_height), "__background__")
        typography = {}
        for item in data["elements"]:
            box = tuple(round(v * scale) for v in item["box"])
            if item["kind"] == "text":
                typography[item["id"]] = add_text(slide, item, box, scale, source, textless)
            elif item["kind"] == "image":
                if item["id"] not in raster_by_id:
                    raise ValueError(f"{data['id']}/{item['id']}: no visible source pixels. Check its box, mask and layer order.")
                picture(slide, raster_by_id[item["id"]], box, item["id"])
            else:
                native_shape(slide, item, box, scale)
        results.append({"id": data["id"], "raster": raster_stats, "typography": typography})

    output.parent.mkdir(parents=True, exist_ok=True)
    file_descriptor, filename = tempfile.mkstemp(suffix=".pptx", dir=output.parent)
    os.close(file_descriptor)
    temporary = Path(filename)
    try:
        deck.save(temporary)
        temporary.replace(output)
    finally:
        temporary.unlink(missing_ok=True)
    report = output.with_suffix(".build.json")
    write_json(report, {"status": "built_pending_render_review", "slides": results})
    return report


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    init = commands.add_parser("init", help="freeze source PNGs and write an empty inventory")
    init.add_argument("job_dir", type=Path)
    init.add_argument("images", type=Path, nargs="+")
    init.add_argument("--slide-width-in", type=float, default=13.33333333)
    build_parser = commands.add_parser("build", help="build the PPTX")
    build_parser.add_argument("manifest", type=Path)
    build_parser.add_argument("output", type=Path)
    build_parser.add_argument("--layers-dir", type=Path, help="write optional source-pixel asset previews")
    args = parser.parse_args()
    try:
        if args.command == "init":
            print(initialize(args.job_dir, args.images, args.slide_width_in))
        else:
            if args.output.suffix.lower() != ".pptx":
                raise ValueError("output filename must end in .pptx")
            print(build(args.manifest, args.output, args.layers_dir))
    except (ValueError, OSError, KeyError) as exc:
        print(f"Reconstruction stopped: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
