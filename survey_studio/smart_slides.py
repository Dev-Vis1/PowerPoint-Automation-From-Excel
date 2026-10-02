"""Build editable survey slides from an Excel workbook, optionally planned by local AI."""

import argparse
import json
from collections import Counter
from pathlib import Path
from urllib.error import URLError
from urllib.request import Request, urlopen

from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

from .survey_to_slides import CORAL, INK, MUTED, PALE, TEAL, WHITE, make_sample, text


DEFAULTS = {
    "Output name": "survey_smart_results.pptx",
    "Title": "Survey results",
    "First slide": "Executive summary",
    "Audience": "Team leads",
    "Data sheet": "Responses",
    "Template": "",
}


def load_survey(path, sheet_name=None):
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        settings = DEFAULTS.copy()
        if "Presentation" in workbook:
            for key, value in workbook["Presentation"].iter_rows(min_row=2, max_col=2, values_only=True):
                if key in settings and value is not None:
                    settings[key] = str(value).strip()
        sheet_name = sheet_name or settings["Data sheet"]
        settings["Data sheet"] = sheet_name
        if sheet_name not in workbook or sheet_name == "Presentation":
            raise ValueError(f"Data sheet '{sheet_name}' does not exist.")
        rows = workbook[sheet_name].iter_rows(values_only=True)
        first_row = next(rows, None)
        if first_row is None:
            raise ValueError("Data sheet is empty; put column headers in the first row.")
        headers = [str(value).strip() if value is not None else "" for value in first_row]
        if not any(headers) or len(set(filter(None, headers))) != len(list(filter(None, headers))):
            raise ValueError("Data sheet needs unique, nonempty column headers.")
        records = []
        for row in rows:
            if any(value is not None and str(value).strip() for value in row):
                records.append({header: row[index] for index, header in enumerate(headers) if header})
            if len(records) > 50000:
                raise ValueError("Limit is 50,000 survey responses per workbook.")
        if not records:
            raise ValueError("Data sheet has no responses.")
        return settings, headers, records
    finally:
        workbook.close()


def profile(headers, records):
    charts = {}
    quotes = []
    for header in headers:
        if not header or any(word in header.lower() for word in ("id", "email", "name", "phone")):
            continue
        values = [str(row[header]).strip() for row in records if row.get(header) is not None and str(row[header]).strip()]
        if not values:
            continue
        counts = Counter(values)
        if len(values) >= 3 and max(len(value) for value in values) >= 35 and len(counts) / len(values) > 0.8:
            quotes.append({"column": header, "samples": values[:3]})
            continue
        if not 2 <= len(counts) <= 12 or len(values) < 3 or len(counts) / len(values) > 0.8:
            continue
        if max(len(value) for value in counts) > 38:
            continue
        ordered = counts.most_common(7)
        if len(counts) > 7:
            ordered.append(("Other", sum(counts.values()) - sum(count for _, count in ordered)))
        chart_id = f"q{len(charts) + 1}"
        charts[chart_id] = {"column": header, "counts": dict(ordered), "answered": len(values)}
    return charts, quotes


def default_plan(charts, quotes):
    return {"charts": [{"id": chart_id, "type": "bar"} for chart_id in list(charts)[:4]],
            "quotes": bool(quotes)}


def local_ai_plan(charts, quotes, settings, model):
    catalog = [{"id": chart_id, "question": chart["column"], "counts": chart["counts"]}
               for chart_id, chart in charts.items()]
    instruction = (
        "You are choosing an executive survey presentation outline. Respond with ONLY JSON: "
        '{"charts":[{"id":"q1","type":"bar"}],"quotes":true}. '
        "Choose 1-4 distinct chart IDs from the catalog (or zero if none). Order the most useful "
        "questions first. Type must be bar or column; use bar for long labels. "
        "Choose quotes only if available. No extra keys, headings, statistics, or invented values. "
        "The first slide is fixed by the workbook and is not part of this JSON.\n"
        + json.dumps({"audience": settings["Audience"], "catalog": catalog,
                      "quotes_available": bool(quotes)}, ensure_ascii=True)
    )
    body = json.dumps({"model": model, "prompt": instruction, "format": "json", "stream": False,
                       "options": {"temperature": 0}}).encode("utf-8")
    request = Request("http://127.0.0.1:11434/api/generate", data=body,
                      headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urlopen(request, timeout=180) as response:
            result = json.load(response)
        return json.loads(result["response"])
    except (URLError, TimeoutError, KeyError, json.JSONDecodeError) as error:
        raise ValueError(f"Local Ollama planning failed: {error}. Start Ollama and check 'ollama list'.") from error


def validate_plan(plan, charts, quotes):
    if not isinstance(plan, dict) or set(plan) != {"charts", "quotes"}:
        raise ValueError("Plan must contain only 'charts' and 'quotes'.")
    selected = plan["charts"]
    if not isinstance(selected, list) or len(selected) > 4 or not isinstance(plan["quotes"], bool):
        raise ValueError("Plan must select at most four charts and a quotes boolean.")
    seen = set()
    for item in selected:
        if not isinstance(item, dict) or set(item) != {"id", "type"}:
            raise ValueError("Each chart must specify only an id and type.")
        if item["id"] not in charts or item["id"] in seen or item["type"] not in ("bar", "column"):
            raise ValueError(f"Invalid or repeated chart in plan: {item!r}")
        seen.add(item["id"])
    if plan["quotes"] and not quotes:
        raise ValueError("Plan requested quotes, but there are no long-text responses.")
    return plan


def new_slide(presentation, title, index, total, source, templated, backgrounds=None):
    layout = next((layout for layout in presentation.slide_layouts if layout.name.lower() == "blank"),
                  min(presentation.slide_layouts, key=lambda layout: len(layout.placeholders)))
    slide = presentation.slides.add_slide(layout)
    if backgrounds:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = backgrounds[(index - 1) % len(backgrounds)]
    elif not templated:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = PALE
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.11), Inches(7.5))
    accent.fill.solid()
    accent.fill.fore_color.rgb = TEAL
    accent.line.fill.background()
    text(slide, "SURVEY BRIEF", 0.55, 0.30, 5, 0.30, 11, TEAL, True)
    text(slide, title[:90], 0.55, 0.87, 12.1, 0.88, 28, INK, True)
    text(slide, f"Source: {source}"[:78], 0.55, 7.05, 10.5, 0.25, 10, MUTED)
    text(slide, f"{index} / {total}", 12, 7.05, 0.8, 0.25, 10, MUTED)
    return slide


def draw_chart(slide, chart_info, chart_type, total):
    counts = chart_info["counts"]
    answered = chart_info["answered"]
    leading, amount = next(iter(counts.items()))
    text(slide, f"{answered} of {total} answered  |  Most common: {leading} ({amount})"[:115],
         0.58, 1.81, 12.1, 0.47, 15, MUTED)
    chart_data = CategoryChartData()
    chart_data.categories = list(counts)
    chart_data.add_series("Responses", list(counts.values()))
    if chart_type == "column" and (len(counts) > 6 or max(map(len, counts)) > 18):
        chart_type = "bar"
    ppt_type = XL_CHART_TYPE.COLUMN_CLUSTERED if chart_type == "column" else XL_CHART_TYPE.BAR_CLUSTERED
    chart = slide.shapes.add_chart(ppt_type, Inches(0.7), Inches(2.43),
                                   Inches(11.8), Inches(4.25), chart_data).chart
    chart.has_legend = False
    chart.chart_style = 10
    chart.value_axis.minimum_scale = 0
    chart.value_axis.maximum_scale = max(counts.values()) + 2
    chart.value_axis.major_unit = max(1, (max(counts.values()) + 4) // 4)
    chart.plots[0].series[0].format.fill.solid()
    chart.plots[0].series[0].format.fill.fore_color.rgb = CORAL if chart_type == "column" else TEAL
    chart.plots[0].has_data_labels = True
    chart.plots[0].data_labels.position = (XL_LABEL_POSITION.OUTSIDE_END if chart_type == "bar"
                                           else XL_LABEL_POSITION.OUTSIDE_END)


def render(settings, records, charts, quotes, plan, output, template):
    presentation = Presentation(template) if template else Presentation()
    if template and abs(presentation.slide_width / presentation.slide_height - 16 / 9) > 0.02:
        raise ValueError("Template must be a 16:9 PowerPoint presentation.")
    backgrounds = []
    if template:
        for example in presentation.slides:
            if example.background.fill.type is not None:
                try:
                    backgrounds.append(example.background.fill.fore_color.rgb)
                except (AttributeError, TypeError):
                    pass
        for slide_id in list(presentation.slides._sldIdLst):
            presentation.part.drop_rel(slide_id.rId)
            presentation.slides._sldIdLst.remove(slide_id)
    if not template:
        presentation.slide_width = Inches(13.333)
        presentation.slide_height = Inches(7.5)
    total_slides = 1 + len(plan["charts"]) + int(plan["quotes"])
    slide = new_slide(presentation, settings["First slide"], 1, total_slides,
                      settings["Data sheet"], bool(template), backgrounds)
    text(slide, settings["Title"][:110], 0.58, 2.00, 11.6, 0.72, 27, INK, True)
    text(slide, f"{len(records)} responses  |  Prepared for {settings['Audience']}"[:110],
         0.58, 2.87, 11.6, 0.40, 17, MUTED)
    text(slide, "KEY QUESTIONS", 0.58, 4.00, 8, 0.35, 13, TEAL, True)
    if plan["charts"]:
        for position, item in enumerate(plan["charts"][:3]):
            chart = charts[item["id"]]
            leading, count = next(iter(chart["counts"].items()))
            line = f"{chart['column']}: {leading} ({count} of {chart['answered']})"
            text(slide, line[:100], 0.58, 4.52 + position * 0.57, 11.5, 0.48, 16)
    else:
        text(slide, "No comparable categorical questions found.", 0.58, 4.52, 11, 0.48, 16)
    for index, item in enumerate(plan["charts"], start=2):
        chart = charts[item["id"]]
        slide = new_slide(presentation, chart["column"], index, total_slides,
                          settings["Data sheet"], bool(template), backgrounds)
        draw_chart(slide, chart, item["type"], len(records))
    if plan["quotes"]:
        slide = new_slide(presentation, "In their words", total_slides, total_slides,
                          settings["Data sheet"], bool(template), backgrounds)
        source = quotes[0]
        text(slide, source["column"], 0.58, 1.85, 11.5, 0.43, 16, MUTED)
        for position, quote in enumerate(source["samples"]):
            text(slide, f'"{quote[:190]}"', 0.66, 2.56 + position * 1.35,
                 11.6, 1.13, 17, INK)
    presentation.save(output)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=Path("examples/demo_survey.xlsx"))
    parser.add_argument("--output", type=Path, help="Override the workbook output name")
    parser.add_argument("--sample", action="store_true", help="Create a new sample workbook; never overwrite one")
    parser.add_argument("--ai", action="store_true", help="Use local Ollama to select slides")
    parser.add_argument("--model", default="llama3.1:8b", help="Installed Ollama model name")
    parser.add_argument("--plan-only", action="store_true", help="Print the validated plan without writing slides")
    args = parser.parse_args()
    if args.sample:
        if args.input.exists():
            parser.error(f"Refusing to overwrite workbook: {args.input}")
        make_sample(args.input)
    try:
        settings, headers, records = load_survey(args.input)
        charts, quotes = profile(headers, records)
        plan = validate_plan(local_ai_plan(charts, quotes, settings, args.model) if args.ai
                             else default_plan(charts, quotes), charts, quotes)
        if args.plan_only:
            print(json.dumps(plan, indent=2))
            return
        name = settings["Output name"]
        if not name or Path(name).name != name or Path(name).suffix.lower() != ".pptx":
            raise ValueError("Output name must be a .pptx filename with no directory.")
        output = args.output or args.input.parent / name
        if output.resolve() == args.input.resolve():
            raise ValueError("Output presentation cannot replace the input workbook.")
        template = args.input.parent / settings["Template"] if settings["Template"] else None
        render(settings, records, charts, quotes, plan, output, template)
        print(f"Created {output} with {1 + len(plan['charts']) + int(plan['quotes'])} editable slides")
    except (ValueError, FileNotFoundError) as error:
        parser.error(str(error))


if __name__ == "__main__":
    main()