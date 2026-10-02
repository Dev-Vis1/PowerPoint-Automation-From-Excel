import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from pptx import Presentation
from pptx.util import Inches

from survey_studio.smart_slides import load_survey, profile, default_plan, render, validate_plan


class SmartSlidesTests(unittest.TestCase):
    def test_different_survey_schema_creates_editable_chart(self):
        with tempfile.TemporaryDirectory() as folder:
            input_path = Path(folder) / "different.xlsx"
            output_path = Path(folder) / "result.pptx"
            workbook = Workbook()
            controls = workbook.active
            controls.title = "Presentation"
            controls.append(("Setting", "Value"))
            controls.append(("Data sheet", "Product feedback"))
            controls.append(("First slide", "Product review"))
            controls.append(("Title", "Autumn release"))
            data = workbook.create_sheet("Product feedback")
            data.append(("Feature", "Rating", "Thoughts"))
            for feature, rating, comment in (
                ("Search", 4, "Search results could use better sorting options."),
                ("Export", 2, "The export dialog is confusing for new users."),
                ("Search", 5, "Search is now much easier to use every day."),
                ("Export", 3, "I wish the export included field names."),
                ("Search", 4, "The new filters make finding items much faster."),
                ("Export", 2, "File downloads sometimes take a long time."),
            ):
                data.append((feature, rating, comment))
            workbook.save(input_path)

            settings, headers, records = load_survey(input_path)
            charts, quotes = profile(headers, records)
            plan = validate_plan(default_plan(charts, quotes), charts, quotes)
            render(settings, records, charts, quotes, plan, output_path, None)

            deck = Presentation(output_path)
            self.assertEqual(settings["First slide"], "Product review")
            self.assertEqual(len(deck.slides), 4)
            self.assertEqual(sum(shape.has_chart for slide in deck.slides for shape in slide.shapes), 2)
            self.assertEqual(sum(shape.shape_type == 13 for slide in deck.slides for shape in slide.shapes), 0)

    def test_unknown_chart_is_rejected(self):
        with self.assertRaisesRegex(ValueError, "Invalid or repeated"):
            validate_plan({"charts": [{"id": "made-up", "type": "bar"}], "quotes": False}, {}, [])

    def test_empty_widescreen_template_renders_native_chart(self):
        with tempfile.TemporaryDirectory() as folder:
            template = Path(folder) / "template.pptx"
            output = Path(folder) / "out.pptx"
            source = Presentation()
            source.slide_width = Inches(13.333)
            source.slide_height = Inches(7.5)
            source.save(template)
            settings = {"First slide": "Overview", "Title": "Survey", "Audience": "Team",
                        "Data sheet": "Answers"}
            records = [{"Answer": "Yes"}, {"Answer": "No"}, {"Answer": "Yes"}]
            charts, quotes = profile(["Answer"], records)
            render(settings, records, charts, quotes, default_plan(charts, quotes), output, template)
            deck = Presentation(output)
            self.assertEqual(len(deck.slides), 2)
            self.assertEqual(sum(shape.has_chart for slide in deck.slides for shape in slide.shapes), 1)


if __name__ == "__main__":
    unittest.main()