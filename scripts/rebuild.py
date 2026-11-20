"""Freeze source slides and assemble an editable PowerPoint from a scene manifest."""

import argparse
import hashlib
import io
import json
import math
import os
import re
import shutil
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


def check_color(obj, key, label):
    if key in obj and (not isinstance(obj[key], str) or not HEX.fullmatch(obj[key])):
        raise ValueError(f"{label}: {key} must be #RRGGBB")


def check_completion(item, root, label):
    choice(item, "occlusion", ("none",), label)
    if "completion" not in item:
        return
    completion = item["completion"]
    label += "/completion"
    if not isinstance(completion, dict) or completion.get("mode") not in ("rgba", "solid"):
        raise ValueError(f"{label}: expected a completion object with mode rgba or solid")
    if completion["mode"] == "rgba":
        unsupported = completion.keys() - {"mode", "file", "origin", "visible", "mask"}
        if unsupported:
            raise ValueError(f"{label}: unsupported rgba fields: {', '.join(sorted(unsupported))}")
        if completion.get("origin") not in ("original", "reconstructed"):
            raise ValueError(f"{label}: rgba origin must be original or reconstructed")
        choice(completion, "visible", ("source", "asset"), label)
        source_path(root, completion.get("file"))
        if "mask" in completion:
            source_path(root, completion["mask"])
        return

    unsupported = completion.keys() - {"mode", "geometry", "mask", "polygon", "color", "radius_px"}
    if unsupported:
        raise ValueError(f"{label}: unsupported solid fields: {', '.join(sorted(unsupported))}")
    silhouette = completion.keys() & {"geometry", "mask", "polygon"}
    if len(silhouette) != 1:
        raise ValueError(f"{label}: solid completion requires exactly one of geometry, mask, or polygon")
    choice(completion, "geometry", tuple(SHAPES), label)
    check_color(completion, "color", label)
    box = item["box"]
    if "mask" in completion:
        source_path(root, completion["mask"])
    if "polygon" in completion:
        check_polygon(completion["polygon"], box, f"{label}/polygon")
    if "radius_px" in completion:
        if completion.get("geometry") != "round_rect":
            raise ValueError(f"{label}: radius_px is only available for round_rect")
        number(completion["radius_px"], 0, min(box[2:]) / 2, f"{label}/radius_px")


def matching_aspect(width, height, reference, label):
    base_width, base_height = reference
    if width * base_height != height * base_width:
        raise ValueError(f"{label}: mixed aspect ratios. Pad or crop all sources to the intended ratio before init.")


def load_scene(path):
    scene = json.loads(path.read_text(encoding="utf-8-sig"))
    if (not isinstance(scene, dict) or type(scene.get("version")) is not int or scene["version"] != 1
            or not isinstance(scene.get("slides"), list) or not scene["slides"]):
        raise ValueError("expected a version 1 manifest with one or more slides")
    number(scene.get("slide_width_in", 13.33333333), 2, 56, "slide_width_in")
    known = set()
    sources = []
    reference = None
    for index, slide in enumerate(scene["slides"], 1):
        if not isinstance(slide, dict):
            raise ValueError(f"slide {index}: expected an object")
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
        if reference is None:
            reference = width, height
        matching_aspect(width, height, reference, slide_id)
        elements = slide.get("elements")
        if not isinstance(elements, list) or not elements:
            raise ValueError(f"{slide_id}: element inventory is empty")
        item_ids = set()
        for index, item in enumerate(elements, 1):
            if not isinstance(item, dict):
                raise ValueError(f"{slide_id}/element {index}: expected an object")
            item_id = item.get("id")
            check_id(item_id, item_ids, slide_id)
            label = f"{slide_id}/{item_id}"
            kind = item.get("kind")
            if kind not in ("text", "image", "shape"):
                raise ValueError(f"{label}: unknown kind: {kind!r}")
            box = item.get("box")
            check_box(box, width, height, label)
            for key in ("color", "erase_fill", "underlay_fill", "fill", "stroke"):
                check_color(item, key, label)
            if kind == "image":
                check_completion(item, path.parent, label)
            elif "completion" in item or "occlusion" in item:
                raise ValueError(f"{label}: completion and occlusion are only available for image elements")
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
                    if "polygon" in item:
                        raise ValueError(f"{label}: choose mask or polygon")
                if "underlay_patch" in item:
                    source_path(path.parent, item["underlay_patch"])
                    if "underlay_fill" in item:
                        raise ValueError(f"{label}: choose underlay_patch or underlay_fill")
                if kind == "shape":
                    if item.get("shape") not in tuple(SHAPES) or "fill" not in item:
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
    if not inputs:
        raise ValueError("provide at least one source image")
    job_dir.mkdir(parents=True, exist_ok=True)
    manifest = job_dir / "manifest.json"
    if manifest.exists():
        raise ValueError(f"manifest already exists: {manifest}")
    sources = [job_dir / f"s{i:03}.png" for i in range(1, len(inputs) + 1)]
    for target in sources:
        if target.exists():
            raise ValueError(f"refusing to overwrite frozen source: {target}")

    with tempfile.TemporaryDirectory(prefix=".ppt-init-", dir=job_dir) as staged_dir:
        staging = Path(staged_dir)
        slides = []
        reference = None
        for path, target in zip(inputs, sources):
            with Image.open(path) as raw:
                source = ImageOps.exif_transpose(raw).convert("RGBA")
            if reference is None:
                reference = source.size
            matching_aspect(*source.size, reference, path)
            plate = Image.new("RGBA", source.size, "white")
            plate.alpha_composite(source)
            frozen = staging / target.name
            plate.convert("RGB").save(frozen, format="PNG")
            slides.append({"id": target.stem, "source": target.name, "width": source.width,
                           "height": source.height, "sha256": digest(frozen), "elements": []})
        write_json(staging / manifest.name, {"version": 1, "slide_width_in": slide_width, "slides": slides})

        # Publish the manifest last. Exclusive writes avoid overwriting another
        # process or user file, and a failed publication removes only its own files.
        created = []
        published = False
        try:
            for target in [*sources, manifest]:
                with (staging / target.name).open("rb") as frozen, target.open("xb") as output:
                    created.append(target)
                    shutil.copyfileobj(frozen, output)
            published = True
        finally:
            if not published:
                for target in reversed(created):
                    target.unlink(missing_ok=True)
    return manifest


def build(manifest, output, layers_dir=None, *, require_complete=False):
    manifest = manifest.resolve()
    scene, sources = load_scene(manifest)
    if layers_dir:
        slide_ids = [slide["id"].casefold() for slide in scene["slides"]]
        if len(set(slide_ids)) != len(slide_ids):
            raise ValueError("slide IDs that differ only by case conflict in the layer-preview directory")
        for data in scene["slides"]:
            image_ids = [item["id"].casefold() for item in data["elements"] if item["kind"] == "image"]
            if len(set(image_ids)) != len(image_ids):
                raise ValueError(f"{data['id']}: image IDs that differ only by case conflict in the layer-preview directory")
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
        unresolved = raster_stats.get("unresolved_images", [])
        if require_complete and unresolved:
            raise ValueError(f"{data['id']}: incomplete images: {', '.join(unresolved)}. Supply complete artwork or a supported completion.")
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
    build_parser.add_argument("--require-complete", action="store_true", help="refuse unresolved image occlusions")
    args = parser.parse_args()
    try:
        if args.command == "init":
            print(initialize(args.job_dir, args.images, args.slide_width_in))
        else:
            if args.output.suffix.lower() != ".pptx":
                raise ValueError("output filename must end in .pptx")
            print(build(args.manifest, args.output, args.layers_dir, require_complete=args.require_complete))
    except (ValueError, OSError, KeyError) as exc:
        print(f"Reconstruction stopped: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
