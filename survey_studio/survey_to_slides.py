"""Create a sample survey workbook and an editable PowerPoint report."""

import argparse
from collections import Counter
from pathlib import Path

from openpyxl import Workbook, load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt


HEADERS = (
    "Response ID",
    "Department",
    "Satisfaction (1-5)",
    "Recommend (0-10)",
    "Improvement Area",
    "Comment",
)
SAMPLE_RESPONSES = (
    (1, "Sales", 5, 10, "Reporting", "The new dashboard saves me time."),
    (2, "Support", 4, 8, "Training", "More examples in training would help."),
    (3, "Operations", 3, 7, "Speed", "Pages take a while to load."),
    (4, "Sales", 4, 9, "Reporting", "Exporting reports could be easier."),
    (5, "Finance", 5, 10, "Integrations", "Works well with our monthly workflow."),
    (6, "Support", 2, 5, "Speed", "Searching takes too long."),
    (7, "Operations", 4, 8, "Training", "A quick-start guide would be useful."),
    (8, "Sales", 5, 10, "Mobile access", "I can check results on the go."),
    (9, "Finance", 3, 6, "Reporting", "Need more flexible filters."),
    (10, "Support", 4, 9, "Integrations", "Customer history is easier to find."),
    (11, "Operations", 2, 4, "Speed", "Loading delays interrupt my work."),
    (12, "Sales", 4, 8, "Mobile access", "Mobile layout could be simpler."),
    (13, "Finance", 5, 9, "Reporting", "The charts are easy to share."),
    (14, "Support", 3, 7, "Training", "Please add role-specific tutorials."),
    (15, "Operations", 4, 8, "Integrations", "Fewer manual handoffs now."),
    (16, "Sales", 5, 10, "Reporting", "Weekly reporting is much faster."),
    (17, "Finance", 3, 7, "Speed", "Large reports are slow."),
    (18, "Support", 4, 8, "Training", "A searchable help page would help."),
    (19, "Operations", 4, 9, "Mobile access", "Useful when away from my desk."),
    (20, "Sales", 2, 5, "Speed", "Some screens freeze briefly."),
    (21, "Finance", 4, 8, "Reporting", "Totals are clear and consistent."),
    (22, "Support", 5, 10, "Integrations", "One place for all customer details."),
    (23, "Operations", 3, 6, "Training", "More onboarding time would help."),
    (24, "Sales", 4, 9, "Reporting", "Good overview of the pipeline."),
)

INK = RGBColor(29, 48, 53)
TEAL = RGBColor(16, 115, 112)
CORAL = RGBColor(223, 111, 83)
PALE = RGBColor(245, 248, 247)
MUTED = RGBColor(95, 112, 114)
WHITE = RGBColor(255, 255, 255)


def make_sample(path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Responses"
    sheet.append(HEADERS)
    for response in SAMPLE_RESPONSES:
        sheet.append(response)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column, width in {"A": 16, "B": 18, "C": 23, "D": 21, "E": 22, "F": 48}.items():
        sheet.column_dimensions[column].width = width
    settings = workbook.create_sheet("Presentation")
    settings.append(("Setting", "Value"))
    for key, value in (
        ("Output name", "survey_smart_results.pptx"),
        ("Title", "Employee survey results"),
        ("First slide", "Executive summary"),
        ("Audience", "Team leads"),
        ("Data sheet", "Responses"),
        ("Template", ""),
    ):
        settings.append((key, value))
    settings.column_dimensions["A"].width = 24
    settings.column_dimensions["B"].width = 42
    workbook.save(path)


def read_responses(path):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if "Responses" not in workbook:
            raise ValueError("Workbook needs a sheet named 'Responses'.")
        rows = workbook["Responses"].iter_rows(values_only=True)
        headers = next(rows, None)
        if headers is None or any(header not in headers for header in HEADERS):
            raise ValueError(f"Responses sheet needs columns: {', '.join(HEADERS)}")
        positions = {header: headers.index(header) for header in HEADERS}
        responses = []
        for row_number, row in enumerate(rows, start=2):
            if not any(value is not None for value in row):
                continue
            response = {header: row[index] for header, index in positions.items()}
            satisfaction = response["Satisfaction (1-5)"]
            recommend = response["Recommend (0-10)"]
            if type(satisfaction) not in (int, float) or satisfaction not in range(1, 6):
                raise ValueError(f"Row {row_number}: satisfaction must be an integer from 1 to 5.")
            if type(recommend) not in (int, float) or recommend not in range(11):
                raise ValueError(f"Row {row_number}: recommend must be an integer from 0 to 10.")
            if not response["Department"] or not response["Improvement Area"]:
                raise ValueError(f"Row {row_number}: department and improvement area are required.")
            responses.append(response)
        if not responses:
            raise ValueError("Responses sheet has no survey responses.")
        return responses
    finally:
        workbook.close()


def text(slide, value, x, y, width, height, size=16, color=INK, bold=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(width), Inches(height))
    frame = box.text_frame
    frame.word_wrap = True
    frame.margin_left = Inches(0.02)
    frame.margin_top = Inches(0.02)
    paragraph = frame.paragraphs[0]
    paragraph.text = str(value)
    paragraph.font.name = "Aptos"
    paragraph.font.size = Pt(size)
    paragraph.font.bold = bold
    paragraph.font.color.rgb = color
    return box


def base_slide(presentation, title, index):
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = PALE
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.12), Inches(7.5))
    accent.fill.solid()
    accent.fill.fore_color.rgb = TEAL
    accent.line.fill.background()
    text(slide, "SURVEY SNAPSHOT", 0.58, 0.35, 5, 0.3, 11, TEAL, True)
    text(slide, title, 0.58, 0.88, 12, 0.65, 29, INK, True)
    text(slide, "Source: Responses worksheet", 0.58, 7.08, 8, 0.24, 10, MUTED)
    text(slide, f"{index} / 3", 12.0, 7.08, 0.7, 0.24, 10, MUTED)
    return slide


def metric(slide, label, value, x):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(1.95), Inches(3.75), Inches(1.55))
    shape.fill.solid()
    shape.fill.fore_color.rgb = WHITE
    shape.line.fill.background()
    text(slide, value, x + 0.25, 2.12, 3.2, 0.64, 30, TEAL, True)
    text(slide, label, x + 0.25, 2.87, 3.2, 0.36, 14, MUTED)


def bar_chart(slide, counts, x, y, width, height, color, max_value):
    data = CategoryChartData()
    data.categories = list(counts)
    data.add_series("Responses", list(counts.values()))
    chart = slide.shapes.add_chart(
        XL_CHART_TYPE.BAR_CLUSTERED,
        Inches(x), Inches(y), Inches(width), Inches(height), data,
    ).chart
    chart.has_legend = False
    chart.chart_style = 10
    chart.value_axis.minimum_scale = 0
    chart.value_axis.maximum_scale = max_value
    chart.value_axis.major_unit = max(1, (max_value + 3) // 4)
    chart.plots[0].series[0].format.fill.solid()
    chart.plots[0].series[0].format.fill.fore_color.rgb = color
    chart.plots[0].has_data_labels = True
    chart.plots[0].data_labels.position = XL_LABEL_POSITION.OUTSIDE_END


def make_presentation(responses, output):
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    total = len(responses)
    satisfaction = Counter(int(row["Satisfaction (1-5)"]) for row in responses)
    priorities = Counter(str(row["Improvement Area"]).strip() for row in responses)
    departments = Counter(str(row["Department"]).strip() for row in responses)
    recommend_rate = sum(row["Recommend (0-10)"] >= 8 for row in responses) / total

    slide = base_slide(presentation, "How respondents feel", 1)
    metric(slide, "Responses", str(total), 0.58)
    metric(slide, "Average satisfaction / 5", f"{sum(score * count for score, count in satisfaction.items()) / total:.1f}", 4.54)
    metric(slide, "Recommend score 8-10", f"{recommend_rate:.0%}", 8.5)
    text(slide, "PARTICIPATION BY DEPARTMENT", 0.58, 4.15, 7, 0.32, 13, TEAL, True)
    for position, (department, count) in enumerate(departments.most_common()):
        top = 4.65 + position * 0.43
        text(slide, department, 0.58, top, 5, 0.32, 17)
        text(slide, str(count), 5.7, top, 1, 0.32, 17, INK, True)

    slide = base_slide(presentation, "Satisfaction distribution", 2)
    text(slide, "Number of responses at each score (1 = low, 5 = high)", 0.58, 1.63, 11, 0.4, 16, MUTED)
    scores = {str(score): satisfaction[score] for score in range(5, 0, -1)}
    bar_chart(slide, scores, 0.58, 2.15, 11.9, 4.55, TEAL, max(scores.values()) + 2)

    slide = base_slide(presentation, "Where to improve", 3)
    text(slide, "Selected improvement area", 0.58, 1.63, 7, 0.4, 16, MUTED)
    bar_chart(slide, dict(priorities.most_common()), 0.58, 2.12, 7.4, 4.7, CORAL, max(priorities.values()) + 2)
    text(slide, "IN THEIR WORDS", 8.35, 2.14, 4.1, 0.38, 13, TEAL, True)
    comments = [str(row["Comment"]).strip() for row in responses if row["Comment"]]
    for position, comment in enumerate(comments[:3]):
        text(slide, f'"{comment}"', 8.35, 2.78 + position * 1.14, 4.15, 0.85, 16)

    presentation.save(output)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=Path("examples/survey_responses.xlsx"))
    parser.add_argument("--output", type=Path, default=Path("examples/survey_results.pptx"))
    parser.add_argument("--sample", action="store_true", help="Create a sample workbook before building slides")
    args = parser.parse_args()
    if args.sample:
        if args.input.exists():
            parser.error(f"Refusing to overwrite existing workbook: {args.input}")
        make_sample(args.input)
    responses = read_responses(args.input)
    make_presentation(responses, args.output)
    print(f"Created {args.output} from {len(responses)} responses in {args.input}")


if __name__ == "__main__":
    main()