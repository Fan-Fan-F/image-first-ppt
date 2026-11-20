"""Regression tests for objects that look correct only while overlapped."""

import io
import json
import sys
import tempfile
import unittest
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw
from pptx import Presentation


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from check_layers import build_review
from raster import prepare
from rebuild import build, initialize, load_scene


class OverlappingLayersTest(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.back_box = [12, 12, 80, 50]
        self.cover_box = [42, 5, 16, 65]
        self.bg = "#F5F2EA"
        self.blue = "#367BA8"
        self.cover = "#C45037"
        self.back = {"id": "back", "kind": "image", "box": self.back_box, "underlay_fill": self.bg}
        self.front = {"id": "front", "kind": "image", "box": self.cover_box, "underlay_fill": self.bg}
        self.slide = {"id": "s001", "source": "source.png", "width": 110, "height": 75,
                      "elements": [self.back, self.front]}

    def source(self, artwork):
        source = Image.new("RGBA", (110, 75), self.bg)
        source.alpha_composite(artwork, tuple(self.back_box[:2]))
        source.alpha_composite(Image.new("RGBA", tuple(self.cover_box[2:]), self.cover),
                               tuple(self.cover_box[:2]))
        return source.convert("RGB")

    def compose(self, background, layers):
        result = background.convert("RGBA")
        for item, artwork in layers:
            result.alpha_composite(artwork, tuple(item["box"][:2]))
        return np.asarray(result.convert("RGB"))

    def save_scene(self, source):
        source.save(self.root / "source.png")
        manifest = self.root / "manifest.json"
        manifest.write_text(json.dumps({"version": 1, "slides": [self.slide]}), encoding="utf-8")
        return manifest

    def test_old_gap_and_restored_solid_have_identical_static_render(self):
        original_full = Image.new("RGBA", (80, 50), self.blue)
        source = self.source(original_full)
        manifest = self.save_scene(source)
        ripped = self.root / "old-layout.pptx"
        build(manifest, ripped)
        with self.assertRaisesRegex(ValueError, "back"):
            build(manifest, self.root / "strict-rejected.pptx", require_complete=True)
        self.assertFalse((self.root / "strict-rejected.pptx").exists())
        base, layers, stats, _ = prepare(self.slide, source, self.root)
        np.testing.assert_array_equal(self.compose(base, layers), np.asarray(source))
        self.assertEqual(stats["unresolved_images"], ["back"])
        self.assertTrue(np.all(np.asarray(layers[0][1])[:, 30:46, 3] == 0))

        self.back["completion"] = {"mode": "solid", "geometry": "rect"}
        manifest = self.save_scene(source)
        rebuilt = self.root / "completed.pptx"
        build(manifest, rebuilt, require_complete=True)
        base, layers, stats, _ = prepare(self.slide, source, self.root)
        np.testing.assert_array_equal(self.compose(base, layers), np.asarray(source))
        np.testing.assert_array_equal(np.asarray(layers[0][1]), np.asarray(original_full))
        self.assertEqual(stats["visuals"]["back"]["completion"]["restored_hidden_pixels"], 800)
        self.assertEqual(stats["changed_outside_masks"], 0)
        self.assertEqual(stats["unresolved_images"], [])
        pptx = Presentation(rebuilt)
        with Image.open(io.BytesIO(pptx.slides[0].shapes[1].image.blob)) as embedded:
            np.testing.assert_array_equal(np.asarray(embedded), np.asarray(original_full))
        _, _, reports_issue = build_review(manifest, ripped, self.root / "bad-move-qa", ["s001/back"], "none")
        self.assertTrue(reports_issue, "an old cutout must fail even if the new manifest declares completion")
        _, _, reports_issue = build_review(manifest, rebuilt, self.root / "move-qa", ["s001/back"], "none")
        self.assertFalse(reports_issue)

    def test_original_rgba_preserves_window_and_partial_alpha(self):
        artwork = Image.new("RGBA", (80, 50))
        draw = ImageDraw.Draw(artwork)
        draw.ellipse((0, 0, 79, 49), fill=(42, 101, 63, 150))
        draw.ellipse((2, 2, 77, 47), fill="#23985D")
        draw.ellipse((29, 15, 50, 34), fill=(0, 0, 0, 0))
        draw.line([(21, 8), (58, 8)], fill="#F0CB57", width=2)
        source = self.source(artwork)
        artwork.save(self.root / "complete.png")
        full = Image.new("1", source.size)
        full.paste(artwork.getchannel("A").point(lambda x: 255 if x else 0).convert("1"), (12, 12))
        full.save(self.root / "full.png")
        visible = full.copy()
        ImageDraw.Draw(visible).rectangle((42, 5, 57, 69), fill=0)
        visible.save(self.root / "visible.png")
        self.back.update({"mask": "visible.png", "completion": {
            "mode": "rgba", "file": "complete.png", "origin": "original",
            "visible": "asset", "mask": "full.png",
        }})
        base, layers, stats, _ = prepare(self.slide, source, self.root)
        np.testing.assert_array_equal(self.compose(base, layers), np.asarray(source))
        np.testing.assert_array_equal(np.asarray(layers[0][1]), np.asarray(artwork))
        self.assertGreater(stats["visuals"]["back"]["completion"]["restored_hidden_pixels"], 0)
        self.assertEqual(layers[0][1].getpixel((40, 23))[3], 0)

        draw.line([(40, 5), (40, 13)], fill=(0, 0, 0, 0))
        artwork.save(self.root / "complete.png")
        with self.assertRaisesRegex(ValueError, "independent silhouette"):
            prepare(self.slide, source, self.root)

    def test_solid_is_rejected_for_multicolor_object(self):
        artwork = Image.new("RGBA", (80, 50), self.blue)
        ImageDraw.Draw(artwork).rectangle((0, 0, 30, 49), fill="#213242")
        self.back["completion"] = {"mode": "solid", "geometry": "rect"}
        with self.assertRaisesRegex(ValueError, "do not support a solid fill"):
            prepare(self.slide, self.source(artwork), self.root)

    def test_known_original_can_be_completely_hidden(self):
        self.back["box"] = [44, 15, 5, 7]
        artwork = Image.new("RGBA", (5, 7), self.blue)
        artwork.save(self.root / "known.png")
        self.back["completion"] = {"mode": "rgba", "file": "known.png", "origin": "original"}
        source = Image.new("RGB", (110, 75), self.bg)
        source.paste(self.cover, (42, 5, 58, 70))
        base, layers, stats, _ = prepare(self.slide, source, self.root)
        np.testing.assert_array_equal(self.compose(base, layers), np.asarray(source))
        np.testing.assert_array_equal(np.asarray(layers[0][1]), np.asarray(artwork))
        self.assertEqual(stats["visuals"]["back"]["completion"]["restored_hidden_pixels"], 35)

    def test_init_can_retry_after_bad_second_image(self):
        source = self.root / "source.png"
        self.source(Image.new("RGBA", (80, 50), self.blue)).save(source)
        destination = self.root / "job"
        with self.assertRaises((ValueError, OSError)):
            initialize(destination, [source, self.root / "missing.png"], 13.333)
        manifest = initialize(destination, [source], 13.333)
        self.assertEqual(len(json.loads(manifest.read_text(encoding="utf-8"))["slides"]), 1)

    def test_fixed_line_height_is_an_absolute_length(self):
        self.slide["elements"] = [{"id": "caption", "kind": "text", "box": [1, 1, 90, 68],
                                   "text": "First\nSecond", "font_px": 6, "line_height_px": 10,
                                   "color": "#000000", "erase_fill": self.bg}]
        manifest = self.save_scene(Image.new("RGB", (110, 75), self.bg))
        output = self.root / "text.pptx"
        build(manifest, output)
        pptx = Presentation(output)
        paragraph = pptx.slides[0].shapes[1].text_frame.paragraphs[0]
        self.assertAlmostEqual(paragraph.line_spacing.pt, 10 * pptx.slide_width / (110 * 12700), places=2)

    def test_malformed_sheet_fails_with_a_value_error(self):
        manifest = self.root / "example.json"
        manifest.write_text("[]", encoding="utf-8")
        with self.assertRaises(ValueError):
            load_scene(manifest)


if __name__ == "__main__":
    unittest.main()
