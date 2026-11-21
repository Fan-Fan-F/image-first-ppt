"""Regression tests for chart cache, embedded workbook and native editability."""

from io import BytesIO
import json
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
import unittest
from zipfile import ZipFile

from PIL import Image
import numpy as np
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
from rebuild import build
from verify import verify


class NativeChartVerificationTest(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.pptx = self.root / "chart.pptx"
        self.manifest = self.root / "manifest.json"
        self.source = self.root / "source.png"
        Image.new("RGB", (800, 450), "white").save(self.source)
        self.spec = {
            "id": "signed-values", "kind": "chart", "chart_type": "column", "box": [80, 45, 560, 315],
            "categories": ["甲组", "乙组", "丙组"],
            "series": [{"name": "观测值", "values": [-4.125, 0, 1.23456789012345e-8]}],
            "value_axis": {"min": -6, "max": 2, "major_unit": 2, "number_format": "0.00", "title": "差值"},
            "category_axis_title": "分组",
            "source_note": "Controlled test data for signed numbers and explicit axis bounds.",
        }

    def deck(self, flattened=False):
        deck = Presentation()
        deck.slide_width = Inches(10)
        deck.slide_height = Inches(5.625)
        slide = deck.slides.add_slide(deck.slide_layouts[6])
        background = slide.shapes.add_picture(str(self.source), 0, 0, deck.slide_width, deck.slide_height)
        background.name = "__background__"
        position = [round(x * deck.slide_width / 800) for x in self.spec["box"]]
        if flattened:
            slide.shapes.add_picture(str(self.source), *position).name = self.spec["id"]
            return deck, None
        if self.spec["chart_type"] == "scatter":
            data = XyChartData()
            for row in self.spec["series"]:
                series = data.add_series(row["name"])
                for x, y in zip(row["x_values"], row["values"]):
                    series.add_data_point(x, y)
            native_type = XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS
        else:
            data = CategoryChartData()
            data.categories = self.spec["categories"]
            for row in self.spec["series"]:
                data.add_series(row["name"], row["values"])
            native_type = XL_CHART_TYPE.COLUMN_CLUSTERED
        shape = slide.shapes.add_chart(native_type, *position, data)
        shape.name = self.spec["id"]
        shape._element.nvGraphicFramePr.cNvPr.set("descr", self.spec["source_note"])
        chart = shape.chart
        for axis, options in ((chart.category_axis, self.spec.get("x_axis")), (
            chart.value_axis, self.spec.get("value_axis"),
        )):
            if options:
                for setting, property_name in (("min", "minimum_scale"), ("max", "maximum_scale"),
                                               ("major_unit", "major_unit")):
                    if setting in options:
                        setattr(axis, property_name, options[setting])
                if "number_format" in options:
                    axis.tick_labels.number_format = options["number_format"]
                if "title" in options:
                    axis.has_title = True
                    axis.axis_title.text_frame.text = options["title"]
        if "category_axis_title" in self.spec:
            chart.category_axis.has_title = True
            chart.category_axis.axis_title.text_frame.text = self.spec["category_axis_title"]
        return deck, chart

    def write_manifest(self):
        self.manifest.write_text(json.dumps({
            "version": 1, "slides": [{"id": "s001", "source": "source.png", "width": 800,
                                       "height": 450, "elements": [self.spec]}],
        }, ensure_ascii=False), encoding="utf-8")

    def audit(self, deck, expected_status):
        deck.save(self.pptx)
        self.write_manifest()
        result = verify(self.manifest, self.pptx, self.root / "qa", renderer="none")
        self.assertEqual(result["structural_status"], expected_status, result)
        return result["slides"][0]["objects"][0]

    def test_column_round_trips_signed_values_and_editable_workbook(self):
        deck, _ = self.deck()
        chart = self.audit(deck, "checked")
        self.assertEqual(chart["actual_chart_type"], "COLUMN_CLUSTERED")
        self.assertEqual(chart["chart_workbook_status"], "verified")
        self.assertEqual(chart["actual_chart_series"][0]["values"], self.spec["series"][0]["values"])
        self.assertEqual(chart["actual_chart_series"][0]["categories"], self.spec["categories"])

    def test_scatter_links_actual_x_and_y_ranges_for_each_series(self):
        self.spec.pop("categories")
        self.spec.pop("category_axis_title")
        self.spec.update({
            "chart_type": "scatter",
            "x_axis": {"min": -2, "max": 2, "major_unit": 1, "number_format": "0", "title": "横轴"},
            "series": [
                {"name": "曲线", "x_values": [-1.75, -0.1, 1.25], "values": [0.25, 0.75, 0.5]},
                {"name": "参考点", "x_values": [-1.1], "values": [-1e-10]},
            ],
        })
        deck, chart = self.deck()
        good = self.audit(deck, "checked")
        self.assertEqual(good["chart_workbook_status"], "verified")
        self.assertEqual(good["actual_chart_series"], self.spec["series"])
        x_point = chart.part._element.xpath(".//c:xVal/c:numRef/c:numCache/c:pt/c:v")[0]
        x_point.text = "1.75"
        bad = self.audit(deck, "issues_found")
        self.assertTrue(any("x_values cached data differ" in issue for issue in bad["issues"]), bad)

    def test_tampered_workbook_fails_with_intact_chart_cache(self):
        deck, chart = self.deck()
        workbook = chart.part.chart_workbook.xlsx_part
        patched = BytesIO()
        with ZipFile(BytesIO(workbook.blob)) as source, ZipFile(patched, "w") as target:
            for member in source.infolist():
                data = source.read(member)
                if member.filename == "xl/worksheets/sheet1.xml":
                    self.assertIn(b"<v>-4.125</v>", data)
                    data = data.replace(b"<v>-4.125</v>", b"<v>4.125</v>")
                target.writestr(member, data)
        workbook.blob = patched.getvalue()
        result = self.audit(deck, "issues_found")
        self.assertEqual(result["chart_workbook_status"], "invalid")
        self.assertTrue(any("values Excel workbook data differ" in issue for issue in result["issues"]), result)

    def test_missing_workbook_is_not_counted_as_editable(self):
        deck, chart = self.deck()
        chart_space = chart.part._element
        chart_space.remove(chart_space.find(qn("c:externalData")))
        result = self.audit(deck, "issues_found")
        self.assertEqual(result["chart_workbook_status"], "invalid")
        self.assertTrue(any("no embedded editable Excel workbook" in issue for issue in result["issues"]), result)

    def test_flattened_image_with_chart_id_is_not_counted_as_editable(self):
        deck, _ = self.deck(flattened=True)
        result = self.audit(deck, "issues_found")
        self.assertTrue(any("native editable PowerPoint chart" in issue for issue in result["issues"]), result)

    def test_builder_removes_incorrect_draft_art_before_placing_native_chart(self):
        self.spec.update({"erase_fill": "#FFFFFF", "title": "已核对的数据", "data_labels": True,
                          "show_legend": True})
        self.spec["value_axis"].update({"visible": True, "show_major_gridlines": True, "show_tick_labels": False})
        draft = Image.new("RGB", (800, 450), "#FFFFFF")
        draft.paste("#DB1137", (80, 45, 640, 360))
        draft.putpixel((14, 17), (0, 56, 93))
        draft.save(self.source)
        self.write_manifest()
        build(self.manifest, self.pptx)
        report = verify(self.manifest, self.pptx, self.root / "qa", renderer="none")
        self.assertEqual(report["structural_status"], "checked", report)
        self.assertEqual(report["slides"][0]["objects"][0]["chart_workbook_status"], "verified")
        loaded = Presentation(self.pptx)
        background = next(shape for shape in loaded.slides[0].shapes if shape.name == "__background__")
        with Image.open(BytesIO(background.image.blob)) as background_image:
            self.assertTrue(np.all(np.asarray(background_image.crop((80, 45, 640, 360))) == 255))
            self.assertEqual(background_image.getpixel((14, 17)), (0, 56, 93))


if __name__ == "__main__":
    unittest.main()
