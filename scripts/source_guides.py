"""Export exact-coordinate rulers and overlapping crops of frozen slide images."""

import argparse
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import re

from PIL import Image, ImageDraw, ImageFont


def open_source(slide, root):
    relative = Path(slide["source"])
    path = (root / relative).resolve()
    if relative.is_absolute() or not path.is_relative_to(root):
        raise ValueError(f"{slide['id']}: source must stay within manifest directory")
    data = path.read_bytes()
    digest = sha256(data).hexdigest()
    if slide.get("sha256") and slide["sha256"] != digest:
        raise ValueError(f"{slide['id']}: frozen source hash changed")
    with Image.open(BytesIO(data)) as source:
        if source.format != "PNG" or source.mode != "RGB" or source.getexif().get(274, 1) != 1:
            raise ValueError(f"{slide['id']}: run rebuild.py init to freeze an oriented RGB PNG")
        if source.size != (slide["width"], slide["height"]):
            raise ValueError(f"{slide['id']}: source dimensions do not match manifest")
        return source.copy(), digest


def marks(extent):
    return [*range(0, extent, 25), extent]


def ruler(source):
    # Offset and measured extents are recorded separately. Marks do not cover
    # or resample the original image.
    left, top = 57, 34
    width, height = source.size
    right, bottom = left + width, top + height
    image = Image.new("RGB", (right + 56, bottom + 35), "#F5F7FA")
    image.paste(source, (left, top))
    draw = ImageDraw.Draw(image)
    font = ImageFont.load_default()
    draw.rectangle((left - 1, top - 1, right, bottom), outline="#788897")
    for axis, extent in (("x", width), ("y", height)):
        for coordinate in marks(extent):
            major = coordinate % 100 == 0 or coordinate == extent
            length = 9 if major else 4
            color = "#25384A" if major else "#788897"
            if axis == "x":
                p = left + coordinate
                draw.line((p, top - 1, p, top - 1 - length), fill=color)
                draw.line((p, bottom, p, bottom + length), fill=color)
            else:
                p = top + coordinate
                draw.line((left - 1, p, left - 1 - length, p), fill=color)
                draw.line((right, p, right + length, p), fill=color)
            if not major or 0 < extent - coordinate < 45:
                continue
            label = str(coordinate)
            if axis == "x":
                draw.text((p, top - 18), label, font=font, fill=color, anchor="mm")
                draw.text((p, bottom + 20), label, font=font, fill=color, anchor="mm")
            else:
                draw.text((left - 31, p), label, font=font, fill=color, anchor="mm")
                draw.text((right + 31, p), label, font=font, fill=color, anchor="mm")
    draw.text((7, 5), "px", font=font, fill="#25384A")
    return image, [left, top]


def axis_tiles(extent, maximum, overlap):
    start = 0
    while True:
        end = min(start + maximum, extent)
        yield start, end
        if end == extent:
            return
        start = end - overlap


def save_new(image, path):
    with path.open("xb") as stream:
        image.save(stream, format="PNG")


def export_guides(manifest, output, max_width, max_height, overlap, page=None):
    if not 0 <= overlap < min(max_width, max_height):
        raise ValueError("overlap must be nonnegative and smaller than both tile dimensions")
    scene = json.loads(manifest.read_text(encoding="utf-8-sig"))
    if scene.get("version") != 1 or not isinstance(scene.get("slides"), list):
        raise ValueError("expected a version 1 slide manifest")
    slides = [slide for slide in scene["slides"] if page is None or slide.get("id") == page]
    if not slides:
        raise ValueError(f"no slide to export: {page or 'manifest is empty'}")
    identifiers = [slide["id"] for slide in slides]
    if (any(not isinstance(value, str) or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,63}", value)
            for value in identifiers) or len({ident.lower() for ident in identifiers}) != len(identifiers)):
        raise ValueError("slide IDs must use safe filenames and be unique")

    output.mkdir(parents=True, exist_ok=True)
    index_path = output / "source-guides.json"
    if index_path.exists():
        raise ValueError(f"guides already exist: {index_path}")
    report = {"version": 1, "coordinate_system": "source pixels, top-left origin, right and bottom excluded",
              "max_tile_size": [max_width, max_height], "overlap": overlap, "slides": []}
    for slide in slides:
        source, digest = open_source(slide, manifest.resolve().parent)
        slide_id = slide["id"]
        guides, origin = ruler(source)
        ruler_name = f"{slide_id}-ruler.png"
        save_new(guides, output / ruler_name)
        item = {"id": slide_id, "source": slide["source"], "sha256": digest,
                "size": list(source.size), "ruler": {"file": ruler_name, "source_origin": origin}, "tiles": []}
        for y, bottom in axis_tiles(source.height, max_height, overlap):
            for x, right in axis_tiles(source.width, max_width, overlap):
                width, height = right - x, bottom - y
                name = f"{slide_id}-x{x:04}-y{y:04}-w{width:04}-h{height:04}.png"
                save_new(source.crop((x, y, right, bottom)), output / name)
                item["tiles"].append({"file": name, "box": [x, y, width, height]})
        report["slides"].append(item)

    with index_path.open("x", encoding="utf-8") as stream:
        json.dump(report, stream, ensure_ascii=False, indent=2)
        stream.write("\n")
    return index_path


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manifest", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--page", help="export one slide ID instead of all slides")
    parser.add_argument("--tile-width", type=int, default=768)
    parser.add_argument("--tile-height", type=int, default=512)
    parser.add_argument("--overlap", type=int, default=64)
    args = parser.parse_args()
    try:
        print(export_guides(args.manifest, args.output, args.tile_width, args.tile_height, args.overlap, args.page))
    except (ValueError, OSError, KeyError, TypeError) as exc:
        parser.exit(1, f"Guide export stopped: {exc}\n")


if __name__ == "__main__":
    main()
