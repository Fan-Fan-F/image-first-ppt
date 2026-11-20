"""Sample source text and write editable DrawingML with explicit font families."""

import os
import re
import shutil
import subprocess
from functools import lru_cache
from pathlib import Path

import numpy as np
from PIL import Image, ImageFont
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Pt


EAST_ASIAN = re.compile(r"[\u2e80-\u9fff\uac00-\ud7ff\uf900-\ufaff\uff00-\uffef\U00020000-\U000323af]")
SAMPLES = {
    "Aptos": ("aptos.ttf", "aptos-bold.ttf"),
    "Arial": ("arial.ttf", "arialbd.ttf"),
    "Calibri": ("calibri.ttf", "calibrib.ttf"),
    "Georgia": ("georgia.ttf", "georgiab.ttf"),
    "Times New Roman": ("times.ttf", "timesbd.ttf"),
    "Microsoft YaHei": ("msyh.ttc", "msyhbd.ttc"),
    "DengXian": ("deng.ttf", "dengb.ttf"),
    "SimHei": ("simhei.ttf", None),
    "SimSun": ("simsun.ttc", None),
}
SAMPLE_SIZE = 64


@lru_cache(maxsize=64)
def sample_font(family, bold):
    filename = SAMPLES.get(family, (None, None))[int(bold)]
    if filename:
        path = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts" / filename
        if path.is_file():
            return ImageFont.truetype(str(path), SAMPLE_SIZE)
    command = shutil.which("fc-match")
    if command:
        pattern = f"{family}:style=Bold" if bold else family
        try:
            result = subprocess.run([command, "-f", "%{family}\n%{file}\n%{index}", pattern],
                                    capture_output=True, text=True, timeout=3, check=True)
            names, path, index = result.stdout.splitlines()
            if family.casefold() in {part.strip().casefold() for part in names.split(",")}:
                return ImageFont.truetype(path, SAMPLE_SIZE, index=int(index))
        except (ValueError, OSError, subprocess.SubprocessError):
            return None
    return None


def crop_ink(mask):
    bounds = mask.getbbox()
    return mask.crop(bounds) if bounds else None


def source_ink(item, source, textless):
    x, y, w, h = item.get("erase", item["box"])
    bounds = x, y, x + w, y + h
    original = np.asarray(source.crop(bounds), dtype=np.int16)
    clean = np.asarray(textless.crop(bounds), dtype=np.int16)
    delta = np.abs(original - clean).max(axis=2)
    signal = delta >= 22
    if signal.sum() < 4:
        return None, None
    strong = delta >= np.quantile(delta[signal], 0.75)
    rgb = np.median(original[strong], axis=0).round().astype(int)
    ink = crop_ink(Image.fromarray((signal * 255).astype(np.uint8)))
    return ink, "#" + "".join(f"{v:02X}" for v in rgb)


def font_match(item, ink):
    text = item["text"]
    asian = bool(EAST_ASIAN.search(text))
    explicit = item.get("font_ea" if asian else "font") or item.get("font")
    families = [explicit] if explicit else (
        ["Microsoft YaHei", "DengXian", "SimHei", "SimSun"] if asian
        else ["Arial", "Calibri", "Aptos", "Georgia", "Times New Roman"])
    weights = [item["bold"]] if "bold" in item else [False, True]
    line = max(text.split("\n"), key=len)
    # Size and face measured from ink only for an upright, single visual line.
    if ink is None or "\n" in text or item.get("rotation") or len(line) > 200:
        return explicit, item.get("bold", False), None, "visual_review"
    target = np.asarray(ink.resize((160, 40), Image.Resampling.BILINEAR), dtype=np.float32) / 255
    best = None
    for family in families:
        for bold in weights:
            font = sample_font(family, bold)
            if font is None:
                continue
            glyphs = font.getmask(line, mode="L")
            rendered = crop_ink(Image.frombytes("L", glyphs.size, bytes(glyphs)))
            if not rendered or not rendered.width or not rendered.height:
                continue
            candidate = np.asarray(rendered.resize((160, 40), Image.Resampling.BILINEAR),
                                   dtype=np.float32) / 255
            xscale = ink.width / rendered.width
            yscale = ink.height / rendered.height
            aspect_penalty = abs(np.log(xscale / yscale))
            score = float(np.mean(np.abs(candidate - target)) + 0.24 * aspect_penalty)
            if best is None or score < best[0]:
                best = score, family, bold, SAMPLE_SIZE * (xscale * yscale) ** 0.5
    if best is None or (not explicit and best[0] > 0.48):
        return explicit, item.get("bold", False), None, "visual_review"
    return best[1], best[2], best[3], "sampled_candidate"


def set_east_asian(run, family):
    properties = run._r.get_or_add_rPr()
    ea = OxmlElement("a:ea")
    ea.set("typeface", family)
    properties.insert_element_before(ea, "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:extLst")


def add_text(slide, item, box, scale, source, textless):
    ink, sampled_color = source_ink(item, source, textless)
    family, bold, estimate_px, method = font_match(item, ink)
    text = item["text"]
    asian = bool(EAST_ASIAN.search(text))
    latin = item.get("font") or (family if family and not asian else "Arial")
    ea = item.get("font_ea") or (family if family and asian else "Microsoft YaHei")
    fallback_px = 0.75 * item["box"][3] / max(1, len(text.split("\n")))
    em_px = item.get("font_px", estimate_px or fallback_px)
    color = item.get("color") or sampled_color or "#111111"

    shape = slide.shapes.add_textbox(*box)
    shape.name = item["id"]
    shape.rotation = item.get("rotation", 0)
    frame = shape.text_frame
    frame.word_wrap = False
    frame.auto_size = MSO_AUTO_SIZE.NONE
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
    frame.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                             "bottom": MSO_ANCHOR.BOTTOM}[item.get("valign", "top")]
    alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}
    for i, line in enumerate(text.split("\n")):
        paragraph = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        paragraph.alignment = alignment[item.get("align", "left")]
        paragraph.space_before = paragraph.space_after = Pt(0)
        paragraph.line_spacing = Emu(round(item["line_height_px"] * scale)) if "line_height_px" in item else 1.0
        if line:
            run = paragraph.add_run()
            run.text = line
            run.font.name = latin
            set_east_asian(run, ea)
            run.font.bold = bold
            run.font.italic = item.get("italic", False)
            run.font.size = max(Pt(1), min(Pt(4000), round(em_px * scale)))
            run.font.color.rgb = RGBColor.from_string(color[1:].upper())
    return {"font": latin, "font_ea": ea, "bold": bold, "font_px": round(em_px, 2), "color": color,
            "font_evidence": "manifest" if "font" in item or "font_ea" in item else method,
            "size_evidence": "manifest" if "font_px" in item else ("sampled_candidate" if estimate_px else "visual_review"),
            "color_evidence": "manifest" if "color" in item else ("sampled" if sampled_color else "visual_review")}
