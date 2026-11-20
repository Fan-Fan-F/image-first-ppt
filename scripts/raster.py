"""Extract visible pixels from the source slide and repair explicit holes."""

from pathlib import Path

import cv2
import numpy as np
from PIL import Image, ImageChops, ImageColor, ImageDraw


def _bounds(box, canvas_size):
    left, top, width, height = box
    right, bottom = left + width, top + height
    if not 0 <= left < right <= canvas_size[0] or not 0 <= top < bottom <= canvas_size[1]:
        raise ValueError(f"region lies outside canvas: {box}")
    return slice(top, bottom), slice(left, right)


def _local_path(root, filename):
    relative = Path(filename)
    path = (root / relative).resolve()
    if relative.is_absolute() or ".." in relative.parts or not path.is_relative_to(root):
        raise ValueError(f"asset path must stay within the manifest directory: {filename}")
    return path


def _load_mask(root, filename, canvas_size):
    with Image.open(_local_path(root, filename)) as image:
        if image.format != "PNG" or image.size != canvas_size:
            raise ValueError(f"mask must be a full-canvas PNG at {canvas_size}: {filename}")
        grayscale = ImageChops.multiply(image.convert("L"), image.convert("RGBA").getchannel("A"))
        values = np.asarray(grayscale)
        if np.any((values != 0) & (values != 255)):
            raise ValueError(f"mask must encode white foreground and black background: {filename}")
        return values == 255


def _load_patch(root, filename, region):
    rows, columns = region
    size = columns.stop - columns.start, rows.stop - rows.start
    with Image.open(_local_path(root, filename)) as image:
        if image.format != "PNG" or image.size != size or image.getexif().get(274, 1) != 1:
            raise ValueError(f"repair patch must be an oriented PNG of exactly {size}: {filename}")
        if image.convert("RGBA").getchannel("A").getextrema() != (255, 255):
            raise ValueError(f"repair patch must be opaque within its local crop: {filename}")
        return np.array(image.convert("RGB"))


def _region_mask(box, canvas_size, root, *, polygon=None, mask_file=None, geometry=None, radius=None):
    region = _bounds(box, canvas_size)
    left, top, width, height = box
    if mask_file:
        if polygon:
            raise ValueError("use either a mask PNG or a polygon")
        full = _load_mask(root, mask_file, canvas_size)
        mask = full[region]
        if np.count_nonzero(full) != np.count_nonzero(mask):
            raise ValueError(f"mask exceeds the declared box: {mask_file}")
    elif polygon or geometry in {"ellipse", "round_rect", "triangle"}:
        raster = Image.new("1", (width, height))
        draw = ImageDraw.Draw(raster)
        outline = (0, 0, width - 1, height - 1)
        if polygon:
            draw.polygon([(x - left, y - top) for x, y in polygon], fill=1)
        elif geometry == "ellipse":
            draw.ellipse(outline, fill=1)
        elif geometry == "round_rect":
            draw.rounded_rectangle(outline, radius=radius if radius is not None else round(min(width, height) / 6), fill=1)
        else:
            draw.polygon([((width - 1) / 2, 0), (width - 1, height - 1), (0, height - 1)], fill=1)
        mask = np.asarray(raster, dtype=bool)
    else:
        mask = np.ones((height, width), dtype=bool)
    if not np.any(mask):
        raise ValueError("empty mask")
    return region, mask


def _element_mask(element, canvas_size, root, *, text=False):
    box = element.get("erase", element["box"]) if text else element["box"]
    try:
        return _region_mask(
            box, canvas_size, root, polygon=element.get("erase_polygon" if text else "polygon"),
            mask_file=element.get("erase_mask" if text else "mask"),
            geometry=element.get("shape") if element["kind"] == "shape" else None,
        )
    except ValueError as error:
        raise ValueError(f"{element['id']}: {error}") from error


def _rgba_candidate(root, filename, size):
    with Image.open(_local_path(root, filename)) as image:
        if image.format != "PNG" or image.size != size or image.getexif().get(274, 1) != 1:
            raise ValueError(f"complete artwork must be an oriented PNG of exactly {size}: {filename}")
        return np.array(image.convert("RGBA"))


def _solid_candidate(element, source, owned, canvas_size, root):
    completion = element["completion"]
    box = element["box"]
    _, silhouette = _region_mask(
        box, canvas_size, root, polygon=completion.get("polygon"), mask_file=completion.get("mask"),
        geometry=completion.get("geometry"), radius=completion.get("radius_px"),
    )
    pixels = source[owned]
    if len(pixels) < 16:
        raise ValueError("a solid completion needs at least 16 visible pixels as color evidence")
    color = completion.get("color")
    rgb = np.array(ImageColor.getrgb(color)) if color else np.rint(np.median(pixels, axis=0)).astype(np.uint8)
    distance = np.max(np.abs(pixels.astype(np.int16) - rgb.astype(np.int16)), axis=1)
    if (distance <= 10).mean() < 0.95:
        raise ValueError("visible pixels do not support a solid fill; supply a registered complete PNG")
    candidate = np.empty((*silhouette.shape, 4), dtype=np.uint8)
    candidate[..., :3] = rgb
    candidate[..., 3] = silhouette.astype(np.uint8) * 255
    return candidate


def _complete_asset(element, source, owned, mask, higher, occlusion_suspected, canvas_size, root):
    completion = element["completion"]
    mode = completion["mode"]
    visible = completion.get("visible", "source")
    if mode == "rgba":
        candidate = _rgba_candidate(root, completion["file"], (owned.shape[1], owned.shape[0]))
        origin = completion["origin"]
    else:
        candidate = _solid_candidate(element, source, owned, canvas_size, root)
        origin = "solid_geometry"
    support = candidate[..., 3] > 0
    if mode == "rgba" and "mask" in completion:
        _, silhouette = _region_mask(element["box"], canvas_size, root, mask_file=completion["mask"])
        missing_silhouette = int((silhouette & ~support).sum())
        extra_silhouette = int((support & ~silhouette).sum())
        if missing_silhouette or extra_silhouette:
            raise ValueError(f"complete PNG disagrees with its independent silhouette: "
                             f"{missing_silhouette} missing pixels, {extra_silhouette} extra")
    missing_source = int((owned & ~support).sum())
    if missing_source:
        raise ValueError(f"complete asset omits {missing_source} visible source pixels; align its alpha with the source mask")
    unexpected = int((support & ~owned & ~higher).sum())
    if unexpected:
        raise ValueError(f"complete asset adds {unexpected} visible pixels outside the source footprint; refine the source mask or align the asset")
    # A declaration drawn as the full silhouette explicitly identifies more
    # occluded pixels. A candidate still has to cover them behind the occluder.
    missing_hidden = int((mask & higher & ~support).sum())
    if missing_hidden:
        raise ValueError(f"complete asset omits {missing_hidden} declared hidden pixels; use a visible-only source mask if those pixels are outside this object")
    hidden = int((support & higher).sum())
    if hidden == 0 and occlusion_suspected:
        raise ValueError("the completed image contains no pixels behind a suspected occluder; review the full silhouette and layer order")
    if visible == "source":
        candidate[owned, :3] = source[owned]
        candidate[owned, 3] = 255
    candidate[~support] = 0
    return candidate, {
        "status": "completed", "mode": mode, "origin": origin, "visible_pixels_from": visible,
        "restored_hidden_pixels": hidden, "complete_alpha_pixels": int(support.sum()),
    }


def _context(region, image_size, padding):
    rows, columns = region
    height, width = image_size
    top, left = max(0, rows.start - padding), max(0, columns.start - padding)
    bottom, right = min(height, rows.stop + padding), min(width, columns.stop + padding)
    outer = slice(top, bottom), slice(left, right)
    inner = slice(rows.start - top, rows.stop - top), slice(columns.start - left, columns.stop - left)
    return outer, inner


def _local_color(rgb, region, mask, excluded, *, force=False):
    # The ring excludes every declared edit, so other text or extracted artwork
    # does not vote for the replacement color. Adjacent untouched pixels do.
    outer, inner = _context(region, rgb.shape[:2], padding=6)
    padded_mask = np.zeros(rgb[outer].shape[:2], dtype=np.uint8)
    padded_mask[inner] = mask
    eligible = ~padded_mask.astype(bool) & ~excluded[outer]
    for diameter, tolerance, required_fraction in [(5, 2, 0.98), (13, 6, 0.95)]:
        ring = cv2.dilate(padded_mask, np.ones((diameter, diameter), dtype=np.uint8)).astype(bool)
        samples = rgb[outer][ring & eligible]
        if not len(samples):
            continue
        color = np.rint(np.median(samples, axis=0)).astype(np.uint8)
        close = np.max(np.abs(samples.astype(np.int16) - color), axis=1) <= tolerance
        if force or (len(samples) >= 12 and np.mean(close) >= required_fraction):
            return color
    return None


def _inpaint(rgb, mask):
    if not np.any(mask):
        return
    if np.all(mask):
        raise ValueError("inpaint needs untouched pixels around the repair mask; narrow the mask or supply explicit fill colors")
    y, x = np.nonzero(mask)
    bounds = slice(y.min(), y.max() + 1), slice(x.min(), x.max() + 1)
    outer, _ = _context(bounds, rgb.shape[:2], padding=8)
    crop = np.ascontiguousarray(rgb[outer])
    pixels = mask[outer]
    repaired = cv2.inpaint(crop, pixels.astype(np.uint8), 4, cv2.INPAINT_TELEA)
    # Assign only masked pixels even when an image library touches its context.
    rgb[outer][pixels] = repaired[pixels]


def _clear_text(elements, source, canvas_size, root):
    cleaned = source.copy()
    union = np.zeros(source.shape[:2], dtype=bool)
    pending = union.copy()
    regions = []
    for element in elements:
        if element["kind"] != "text":
            continue
        region, mask = _element_mask(element, canvas_size, root, text=True)
        union[region] |= mask
        regions.append((element, region, mask))

    claimed = np.zeros_like(union)
    records = {}
    for element, region, mask in reversed(regions):
        owned = mask & ~claimed[region]
        claimed[region] |= mask
        count = int(owned.sum())
        mode = element.get("erase_mode", "auto")
        fill = element.get("erase_fill")
        patch = element.get("erase_patch")
        if not count:
            method = "already_covered"
        elif patch:
            cleaned[region][owned] = _load_patch(root, patch, region)[owned]
            method = "explicit_patch"
        elif fill:
            cleaned[region][owned] = ImageColor.getrgb(fill)
            method = "explicit_fill"
        else:
            color = None if mode == "inpaint" else _local_color(source, region, mask, union, force=mode == "flat")
            if color is not None:
                cleaned[region][owned] = color
                method = "local_flat"
            elif mode == "flat":
                raise ValueError(f"cannot sample a flat erase color for {element['id']}; set erase_fill")
            else:
                pending[region] |= owned
                method = "local_inpaint"
        records[element["id"]] = {"method": method, "pixels": count}

    _inpaint(cleaned, pending)
    return cleaned, union, records


def prepare(slide, source, root):
    """Return repaired plate, movable artwork, stats, and text-cleared source.

    Element array order is back-to-front. Visible ownership explains the
    flattened source. A completed asset can independently extend behind its
    occluders, and its completed support must never expand the background cut.
    """
    root = Path(root).resolve()
    original = np.asarray(source.convert("RGB"))
    canvas_size = source.size
    cleaned, text_mask, text_records = _clear_text(slide["elements"], original, canvas_size, root)
    background = cleaned.copy()
    visual_mask = np.zeros(original.shape[:2], dtype=bool)
    entries = []
    for element in reversed(slide["elements"]):
        if element["kind"] not in {"image", "shape"}:
            continue
        region, mask = _element_mask(element, canvas_size, root)
        higher = visual_mask[region].copy()
        owned = mask & ~higher
        visual_mask[region] |= mask
        entries.append((element, region, mask, owned, higher))

    assets = []
    records = {}
    unresolved = []
    pending = np.zeros_like(visual_mask)
    for element, region, mask, owned, higher in reversed(entries):
        count = int(owned.sum())
        overlap = int(mask.sum()) - count
        layer_completion = None
        declaration = element.get("completion", {})
        known_original = declaration.get("mode") == "rgba" and declaration.get("origin") == "original"
        if element["kind"] == "image" and (count or known_original):
            # A visible-only mask can stop right at the occluder, leaving zero
            # literal intersection. Nearby upper pixels inside the image box
            # then warrant review without inventing an extent.
            margin = cv2.dilate(owned.astype(np.uint8), np.ones((5, 5), dtype=np.uint8))
            touches_higher = bool(np.any(higher & margin.astype(bool)))
            occlusion_suspected = bool(overlap) or (touches_higher and element.get("occlusion") != "none")
            if "completion" in element:
                try:
                    rgba, layer_completion = _complete_asset(element, cleaned[region], owned, mask, higher, occlusion_suspected,
                                                             canvas_size, root)
                except ValueError as error:
                    raise ValueError(f"{element['id']}: {error}") from error
            else:
                rgba = np.zeros((*mask.shape, 4), dtype=np.uint8)
                rgba[owned, :3] = cleaned[region][owned]
                rgba[owned, 3] = 255
                status = "visible_only"
                if overlap:
                    status = "incomplete_overlap"
                elif occlusion_suspected:
                    status = "possible_overlap"
                if status != "visible_only":
                    unresolved.append(element["id"])
                layer_completion = {"status": status, "restored_hidden_pixels": 0,
                                    "complete_alpha_pixels": count, "touches_higher": touches_higher}
            assets.append((element, Image.fromarray(rgba)))
        elif element["kind"] == "shape":
            layer_completion = {"status": "native_geometry"}

        fill = element.get("underlay_fill")
        patch = element.get("underlay_patch")
        if not count:
            method = "fully_occluded"
        elif patch:
            background[region][owned] = _load_patch(root, patch, region)[owned]
            method = "explicit_patch"
        elif fill:
            background[region][owned] = ImageColor.getrgb(fill)
            method = "explicit_fill"
        else:
            color = _local_color(cleaned, region, mask, visual_mask)
            if color is not None:
                background[region][owned] = color
                method = "local_flat"
            else:
                pending[region] |= owned
                method = "local_inpaint"
        records[element["id"]] = {
            "kind": element["kind"], "visible_pixels": count, "occluded_pixels": overlap,
            "overlapping_text_mask_pixels": int((text_mask[region] & owned).sum()),
            "underlay_method": method,
            "completion": layer_completion,
        }

    _inpaint(background, pending)
    edit_mask = text_mask | visual_mask
    difference = np.any(background != original, axis=2)
    stats = {
        "text_mask_pixels": int(text_mask.sum()),
        "visual_mask_pixels": int(visual_mask.sum()),
        "text": text_records,
        "visuals": records,
        "unresolved_images": unresolved,
        "changed_background_pixels": int(difference.sum()),
        "changed_outside_masks": int((difference & ~edit_mask).sum()),
    }
    return Image.fromarray(background), assets, stats, Image.fromarray(cleaned)
