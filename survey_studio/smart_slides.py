"""Build editable survey slides from an Excel workbook, optionally planned by local AI."""

import argparse
import csv
import io
import itertools
import json
import math
import os
import re
from collections import Counter
from pathlib import Path
from statistics import mean, median, pstdev
from urllib.error import URLError
from urllib.request import Request, urlopen

from openpyxl import load_workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt


INK = RGBColor(89, 89, 89)
TEAL = RGBColor(1, 59, 103)
CORAL = RGBColor(1, 59, 103)
PALE = RGBColor(242, 242, 242)
LIGHT_BLUE = RGBColor(219, 227, 237)
MUTED = RGBColor(89, 89, 89)
HEADING_FONT = "Tw Cen MT"
BODY_FONT = "Century Gothic"
BRAND_LOGO = Path(__file__).parent / "assets" / "amo-logo.png"
OLLAMA_BASE_URL = os.environ.get("OLLAMA_BASE_URL", "http://127.0.0.1:11434").rstrip("/")


def text(slide, value, x, y, width, height, size=16, color=INK, bold=False, font_name=BODY_FONT):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(width), Inches(height))
    frame = box.text_frame
    frame.word_wrap = True
    frame.margin_left = Inches(0.02)
    frame.margin_top = Inches(0.02)
    paragraph = frame.paragraphs[0]
    paragraph.text = str(value)
    paragraph.font.name = font_name
    paragraph.font.size = Pt(size)
    paragraph.font.bold = bold
    paragraph.font.color.rgb = color
    return box


DEFAULTS = {
    "Output name": "survey_smart_results.pptx",
    "Title": "Survey results",
    "First slide": "Executive summary",
    "Audience": "Team leads",
    "Data sheet": "Responses",
    "Template": "",
}


def _looks_like_question_row(field_ids, descriptions):
    if not descriptions or len(descriptions) != len(field_ids):
        return False
    id_fields = all(re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", value or "") for value in field_ids)
    described = sum(any(char.isspace() for char in value) for value in descriptions)
    return id_fields and described >= max(1, len(field_ids) // 2)


def _semantic_text(field, label):
    return re.sub(r"[^a-z0-9]+", " ", f"{field} {label}".lower()).strip()


def load_survey(path, sheet_name=None):
    path = Path(path)
    if path.suffix.lower() == ".csv":
        raw = path.read_bytes()
        try:
            content = raw.decode("utf-8-sig")
        except UnicodeDecodeError:
            content = raw.decode("cp1252")
        try:
            delimiter = csv.Sniffer().sniff(content[:8192], delimiters=";,\t,").delimiter
        except csv.Error:
            delimiter = ";" if content.count(";") >= content.count(",") else ","
        rows = [row for row in csv.reader(io.StringIO(content, newline=""), delimiter=delimiter)
                if any(cell.strip() for cell in row)]
        if not rows:
            raise ValueError("CSV file is empty.")
        first = [cell.strip() for cell in rows[0]]
        has_question_row = len(rows) > 2 and _looks_like_question_row(first, rows[1])
        if has_question_row:
            headers = first
            labels = {key: question.strip() for key, question in zip(first, rows[1])}
            data_rows = rows[2:]
        else:
            headers = first
            labels = {key: key for key in first}
            data_rows = rows[1:]
        if len(set(filter(None, headers))) != len(list(filter(None, headers))):
            raise ValueError("CSV needs unique, nonempty field names in its first row.")
        records = []
        for row_number, row in enumerate(data_rows, start=3 if has_question_row else 2):
            if len(row) > len(headers):
                raise ValueError(f"CSV row {row_number} has more fields than the header.")
            row = row + [""] * (len(headers) - len(row))
            if any(value.strip() for value in row):
                records.append({key: value.strip() or None for key, value in zip(headers, row) if key})
            if len(records) > 50000:
                raise ValueError("Limit is 50,000 survey responses per file.")
        if not records:
            raise ValueError("CSV has no response rows.")
        settings = DEFAULTS.copy()
        settings["Data sheet"] = "Data"
        settings["Title"] = path.stem.replace("_", " ").replace("-", " ")
        settings["Source name"] = path.name
        return settings, headers, records, labels

    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        settings = DEFAULTS.copy()
        settings["Source name"] = path.name
        if "Presentation" in workbook:
            for key, value in workbook["Presentation"].iter_rows(min_row=2, max_col=2, values_only=True):
                if key in settings and value is not None:
                    settings[key] = str(value).strip()
        sheet_name = sheet_name or settings["Data sheet"]
        settings["Data sheet"] = sheet_name
        if sheet_name not in workbook or sheet_name == "Presentation":
            raise ValueError(f"Data sheet '{sheet_name}' does not exist.")
        rows = iter(workbook[sheet_name].iter_rows(values_only=True))
        first_row = next(rows, None)
        if first_row is None:
            raise ValueError("Data sheet is empty; put column headers in the first row.")
        headers = [str(value).strip() if value is not None else "" for value in first_row]
        if not any(headers) or len(set(filter(None, headers))) != len(list(filter(None, headers))):
            raise ValueError("Data sheet needs unique, nonempty column headers.")
        possible_descriptions = next(rows, None)
        descriptions = ([str(value).strip() if value is not None else "" for value in possible_descriptions]
                        if possible_descriptions is not None else None)
        if descriptions and _looks_like_question_row(headers, descriptions):
            labels = {header: description for header, description in zip(headers, descriptions) if header}
        else:
            labels = {header: header for header in headers if header}
            if possible_descriptions is not None:
                rows = itertools.chain([possible_descriptions], rows)
        records = []
        for row in rows:
            if any(value is not None and str(value).strip() for value in row):
                records.append({header: row[index] for index, header in enumerate(headers) if header})
            if len(records) > 50000:
                raise ValueError("Limit is 50,000 survey responses per workbook.")
        if not records:
            raise ValueError("Data sheet has no responses.")
        return settings, headers, records, labels
    finally:
        workbook.close()


def _number(value):
    if value is None or isinstance(value, bool):
        return None
    try:
        result = float(str(value).strip().replace(",", "."))
    except (TypeError, ValueError):
        return None
    return result if math.isfinite(result) else None


def _segment_name(category, value):
    number = _number(value)
    key = str(number) if number is not None else str(value).strip()
    return category.get("value_to_bin", {}).get(key, _vehicle_group(category["field"], value))


def _rating_summary(values, scale_min, scale_max):
    top = sum(value >= 9 for value in values) if scale_max == 10 else 0
    bottom = sum(1 <= value <= 4 for value in values) if scale_max == 10 else 0
    return {
        "n": len(values),
        "mean": mean(values),
        "median": median(values),
        "distribution": {str(score): sum(value == score for value in values)
                         for score in range(scale_min, scale_max + 1)},
        "top_count": top,
        "top_share": top / len(values) if values else 0,
        "bottom_count": bottom,
        "bottom_share": bottom / len(values) if values else 0,
    }


def _vehicle_group(field, value):
    normalized = str(value).strip().lower()
    field_name = field.lower()
    if any(token in field_name for token in ("drive", "engine", "propulsion", "powertrain")):
        if any(token in normalized for token in ("electric", "bev", "battery")):
            return "BEV"
        if any(token in normalized for token in ("plug-in", "plug in", "phev")):
            return "PHEV"
        if any(token in normalized for token in ("hybrid", "hev", "gasoline", "petrol", "diesel", "combustion")):
            return "Combustion engine + hybrid"
    return str(value).strip()


def _display_question(question):
    value = re.sub(r"^how\s+(?:would|do)\s+you\s+(?:evaluate|rate)\s+", "", question, flags=re.I)
    value = re.sub(r"^(?:please\s+)?rate\s+", "", value, flags=re.I)
    value = re.sub(r"^(which|what)\s+", "", value, flags=re.I)
    value = re.sub(r"\s+do you\s+.*$", "", value, flags=re.I)
    value = re.sub(r"\s+(?:of|for)\s+this\s+vehicle\??$", "", value, flags=re.I)
    value = re.sub(r"[?!.]+$", "", value).strip()
    value = re.sub(r"^the\s+", "", value, flags=re.I)
    return value[:1].upper() + value[1:] if value else question[:80]


def _display_group(field, question):
    name = field.lower()
    if "brand" in name or "make" in name:
        return "Current brand"
    if "model" in name:
        return "Current model"
    if any(token in name for token in ("drive", "engine", "propulsion", "powertrain")):
        return "Current drive type"
    short = re.sub(r"^(which|what|how)\s+", "", question, flags=re.I)
    short = re.sub(r"\s+(do you|would you|is your|are your)\b.*$", "", short, flags=re.I)
    short = short.rstrip(" ?.!")
    return short[:64] or field


def analyze_survey(headers, records, labels):
    numeric_fields = []
    numeric_distributions = []
    categorical_fields = []
    excluded = re.compile(r"\b(id|identifier|email|phone|case number|uuid|timestamp)\b", re.I)
    for field in headers:
        label = labels.get(field, field)
        field_label = _semantic_text(field, label)
        if not field or excluded.search(field_label):
            continue
        values = [row.get(field) for row in records if row.get(field) not in (None, "")]
        if not values:
            continue
        numbers = [_number(value) for value in values]
        numeric_values = [value for value in numbers if value is not None]
        if len(numeric_values) == len(values):
            low, high = min(numeric_values), max(numeric_values)
            unique = set(numeric_values)
            rating_language = re.search(
                r"\b(rate|rating|score|evaluate|evaluation|satisfaction|likelihood|design|concept)\b",
                field_label, re.I)
            non_rating_language = re.search(
                r"\b(age|income|price|pricing|budget|willingness to pay|intent|intention|buy|purchase|count|number|years)\b",
                field_label, re.I)
            unlabeled_scale = (not non_rating_language and
                               ((low >= 1 and len(unique) >= 4)
                                or (low == 0 and len(unique) >= 6)))
            scale = (low >= 0 and high <= 10 and 3 <= len(unique) <= 11
                     and all(value.is_integer() for value in unique)
                     and (bool(rating_language) or unlabeled_scale))
            if scale:
                scale_min = 0 if low == 0 else 1
                scale_max = 10
                valid = [value for value in numeric_values if scale_min <= value <= scale_max]
                title = label if label != field else field
                numeric_fields.append({"id": f"rating:{field}", "kind": "rating", "field": field, "label": title,
                                       "display_label": _display_question(title),
                                       "n": len(valid), "missing": len(records) - len(valid),
                                       "scale_min": scale_min, "scale_max": scale_max,
                                       "summary": _rating_summary(valid, scale_min, scale_max)})
                continue
            if len(unique) > 1:
                is_integer = all(value.is_integer() for value in numeric_values)
                if is_integer and high - low <= 120:
                    width = max(1, math.ceil((high - low + 1) / 4))
                    bins = {}
                    start = math.floor(low / width) * width
                    while start <= high:
                        end = start + width - 1
                        bin_label = f"{start}-{end}"
                        count = sum(start <= value <= end for value in numeric_values)
                        if count:
                            bins[bin_label] = count
                        start += width
                else:
                    ordered = sorted(numeric_values)
                    bins = {}
                    for index in range(5):
                        start_index = index * len(ordered) // 5
                        end_index = (index + 1) * len(ordered) // 5
                        group = ordered[start_index:end_index]
                        if group:
                            bin_name = f"{group[0]:g}-{group[-1]:g}"
                            bins[bin_name] = bins.get(bin_name, 0) + len(group)
                candidate_id = f"numeric:{field}"
                numeric_item = {"id": candidate_id, "kind": "numeric", "field": field,
                                "label": label if label != field else field,
                                "n": len(numeric_values), "mean": mean(numeric_values),
                                "median": median(numeric_values), "counts": bins}
                numeric_distributions.append(numeric_item)
                if re.search(r"\b(age|income|education|region|household|demographic)\b", field_label, re.I):
                    value_to_bin = {}
                    if is_integer and high - low <= 120:
                        width = max(1, math.ceil((high - low + 1) / 4))
                        start = math.floor(low / width) * width
                        while start <= high:
                            end = start + width - 1
                            bin_name = f"{start}-{end}"
                            for value in numeric_values:
                                if start <= value <= end:
                                    value_to_bin[str(value)] = bin_name
                            start += width
                    else:
                        ordered = sorted(numeric_values)
                        for index in range(5):
                            start_index = index * len(ordered) // 5
                            end_index = (index + 1) * len(ordered) // 5
                            group_values = ordered[start_index:end_index]
                            if group_values:
                                bin_name = f"{group_values[0]:g}-{group_values[-1]:g}"
                                for value in group_values:
                                    value_to_bin[str(value)] = bin_name
                    categorical_fields.append({"field": field, "label": numeric_item["label"],
                                               "counts": Counter(value_to_bin.values()),
                                               "answered": len(value_to_bin),
                                               "value_to_bin": value_to_bin, "derived": True})
                continue
        counts = Counter(str(value).strip() for value in values)
        if 2 <= len(counts) <= 60 and len(counts) / len(values) <= 0.8:
            title = _display_question(label) if label != field else field
            categorical_fields.append({"field": field, "label": title, "counts": counts,
                                       "question": label, "answered": len(values)})

    slides = {}
    insights = {}
    required_ratings = []
    for rating in numeric_fields:
        rating_id = rating["id"]
        rating["kind"] = "rating"
        slides[rating_id] = rating
        required_ratings.append(rating_id)
        summary = rating["summary"]
        insight_id = f"insight:{rating['field']}:overall"
        insights[insight_id] = (
            f"{rating['display_label']}: mean {summary['mean']:.1f}/10; "
            f"{summary['top_share']:.0%} top-box (9-10), "
            f"{summary['bottom_share']:.0%} bottom-box (1-4), n={summary['n']}."
        )
        rating["insight_id"] = insight_id

    for numeric in numeric_distributions:
        slides[numeric["id"]] = numeric

    for category in categorical_fields:
        if category.get("derived"):
            continue
        grouped = Counter(_vehicle_group(category["field"], value)
                          for value in category["counts"].elements())
        ordered = grouped.most_common(7)
        if len(grouped) > 7:
            ordered.append(("Other", sum(count for _, count in grouped.most_common()[7:])))
        candidate_id = f"frequency:{category['field']}"
        slides[candidate_id] = {"id": candidate_id, "kind": "frequency", "field": category["field"],
                               "label": category["label"], "n": category["answered"],
                               "counts": dict(ordered)}

    intent_fields = [category for category in categorical_fields
                     if re.search(r"\b(buy|purchase|intent|intention|consider|reject|acceptance|openness)\b",
                                  _semantic_text(category["field"], category["label"]), re.I)]
    demographic_fields = [category for category in categorical_fields
                          if re.search(r"\b(age|gender|sex|income|education|region|household|occupation)\b",
                                       _semantic_text(category["field"], category["label"]), re.I)]
    for outcome in intent_fields:
        outcomes = list(outcome["counts"])
        for segment in demographic_fields:
            matrix = {str(value).strip(): Counter() for value in segment["counts"]}
            for record in records:
                group = record.get(segment["field"])
                result = record.get(outcome["field"])
                if group not in (None, "") and result not in (None, ""):
                    group_name = _segment_name(segment, group)
                    if group_name in matrix:
                        matrix[group_name][str(result).strip()] += 1
            rows = [{"name": group, "counts": dict(counts), "n": sum(counts.values())}
                    for group, counts in matrix.items() if sum(counts.values()) >= 10]
            if len(rows) < 2:
                continue
            candidate_id = f"crosstab:{outcome['field']}:{segment['field']}"
            slides[candidate_id] = {"id": candidate_id, "kind": "crosstab",
                                   "field": outcome["field"], "label": outcome["label"],
                                   "group_label": segment["label"], "outcomes": outcomes,
                                   "groups": rows}

    subgroup_id_number = 0
    for rating in numeric_fields:
        for category in categorical_fields:
            group_values = {}
            for record in records:
                value = record.get(category["field"])
                score = _number(record.get(rating["field"]))
                if value in (None, "") or score is None:
                    continue
                group = _segment_name(category, value)
                group_values.setdefault(group, []).append(score)
            # Collapse sparse groups instead of presenting unstable, tiny bases.
            larger = {name: scores for name, scores in group_values.items() if len(scores) >= 10}
            smaller = [(name, scores) for name, scores in group_values.items() if len(scores) < 10]
            if smaller:
                larger["Other"] = [score for _, scores in smaller for score in scores]
            if not 2 <= len(larger) <= 12:
                continue
            groups = [{"name": name, **_rating_summary(scores, rating["scale_min"], rating["scale_max"])}
                      for name, scores in larger.items() if scores]
            subgroup_id_number += 1
            candidate_id = f"subgroup:{rating['field']}:{category['field']}"
            spread = max(row["mean"] for row in groups) - min(row["mean"] for row in groups)
            slides[candidate_id] = {"id": candidate_id, "kind": "subgroup", "field": rating["field"],
                                   "label": rating["display_label"], "group_field": category["field"],
                                   "group_label": _display_group(category["field"], category["label"]),
                                   "question": f"How does the {rating['display_label'].lower()} score differ by "
                                               f"{_display_group(category['field'], category['label']).lower()}?",
                                   "groups": groups,
                                   "spread": spread}
            leader = max(groups, key=lambda row: row["mean"])
            lowest = min(groups, key=lambda row: row["mean"])
            insight_id = f"insight:{rating['field']}:{category['field']}"
            insights[insight_id] = (
                f"{rating['display_label']} by {category['label']}: {leader['name']} leads at "
                f"{leader['mean']:.1f}/10 (n={leader['n']}); {lowest['name']} is lowest at "
                f"{lowest['mean']:.1f}/10 (n={lowest['n']})."
            )
            slides[candidate_id]["insight_id"] = insight_id

    text = _semantic_text(" ".join(headers), " ".join(labels.values()))
    coverage = {
        "demographics": bool(re.search(r"\b(age|gender|sex|income|education|region|household|occupation)\b", text)),
        "purchase_intent_or_rejection": bool(re.search(r"\b(buy|purchase|intent|intention|consider|reject|acceptance)\b", text)),
        "pricing_expectations": bool(re.search(r"\b(price|pricing|willingness to pay|budget|afford)\b", text)),
        "emerging_brand_acceptance": bool(re.search(r"\b(emerging brand|new brand|brand acceptance|brand openness)\b", text)),
    }
    return {"n": len(records), "ratings": numeric_fields, "categories": categorical_fields,
            "slides": slides, "insights": insights, "required_ratings": required_ratings,
            "coverage": coverage}


def default_analysis_plan(analysis):
    subgroup_ids = [key for key, slide in analysis["slides"].items() if slide["kind"] == "subgroup"]
    frequency_ids = [key for key, slide in analysis["slides"].items() if slide["kind"] == "frequency"]
    numeric_ids = [key for key, slide in analysis["slides"].items() if slide["kind"] == "numeric"]
    crosstab_ids = [key for key, slide in analysis["slides"].items() if slide["kind"] == "crosstab"]
    subgroup_ids.sort(key=lambda key: analysis["slides"][key]["spread"], reverse=True)
    return {"slide_ids": (analysis["required_ratings"] + subgroup_ids[:3] + crosstab_ids[:2]
                           + numeric_ids[:2] + frequency_ids[:2])[:8],
            "insight_ids": list(analysis["insights"])[:4]}


def ai_analysis_plan(analysis, settings, model):
    catalog = []
    for candidate_id, slide in analysis["slides"].items():
        if slide["kind"] == "rating":
            summary = slide["summary"]
            detail = {"kind": "rating", "question": slide["label"], "n": summary["n"],
                      "mean": round(summary["mean"], 2), "top_share": round(summary["top_share"], 3),
                      "bottom_share": round(summary["bottom_share"], 3)}
        elif slide["kind"] == "subgroup":
            detail = {"kind": "subgroup", "question": slide["label"], "by": slide["group_label"],
                      "spread": round(slide["spread"], 2),
                      "groups": [{"name": group["name"], "n": group["n"], "mean": round(group["mean"], 2)}
                                 for group in slide["groups"]]}
        elif slide["kind"] == "crosstab":
            detail = {"kind": "crosstab", "question": slide["label"], "by": slide["group_label"],
                      "groups": slide["groups"]}
        elif slide["kind"] == "numeric":
            detail = {"kind": "numeric", "question": slide["label"], "n": slide["n"],
                      "mean": round(slide["mean"], 2), "median": round(slide["median"], 2),
                      "ranges": slide["counts"]}
        else:
            detail = {"kind": "frequency", "question": slide["label"], "n": slide["n"],
                      "counts": slide["counts"]}
        catalog.append({"id": candidate_id, **detail})
    prompt = (
        "Plan a concise customer-research presentation. Use only computed candidates "
        "and factual insight strings provided. Always include every required_rating ID. Choose up "
        "to 4 additional subgroup, crosstab, numeric, or frequency slides from the candidate catalog. "
        "Choose up to 4 supplied insight IDs. Do not invent, hallucinate, or recalculate IDs or "
        "figures. If demographics, pricing, purchase rejection/intent, or emerging-brand acceptance "
        "are not in coverage, do not claim findings about them. Return only JSON with keys slide_ids and insight_ids. "
        'Example format: {"slide_ids": ["<candidate_id_1>", "<candidate_id_2>"], "insight_ids": ["<insight_id_1>"]}.\n'
        + json.dumps({"audience": settings.get("Audience", "Decision makers"),
                      "case_count": analysis["n"], "data_coverage": analysis["coverage"],
                      "required_rating_ids": analysis["required_ratings"], "candidates": catalog,
                      "available_insights": analysis["insights"]}, ensure_ascii=True)
    )
    body = json.dumps({"model": model, "prompt": prompt, "format": "json", "stream": False,
                       "options": {"temperature": 0}}).encode("utf-8")
    request = Request(f"{OLLAMA_BASE_URL}/api/generate", data=body,
                      headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urlopen(request, timeout=180) as response:
            plan = json.loads(json.load(response)["response"])
    except (URLError, TimeoutError, KeyError, json.JSONDecodeError):
        # Graceful fallback: if Ollama is absent, offline, or timed out, use deterministic plan
        return default_analysis_plan(analysis)
    return validate_analysis_plan(plan, analysis)


def validate_analysis_plan(plan, analysis, strict=False):
    if not isinstance(plan, dict) or ("slide_ids" not in plan and "slides" not in plan):
        raise ValueError("AI plan must be a JSON object with 'slide_ids' and 'insight_ids'.")
    raw_slides = plan.get("slide_ids", plan.get("slides", []))
    raw_insights = plan.get("insight_ids", plan.get("insights", []))
    if not isinstance(raw_slides, list) or not isinstance(raw_insights, list):
        raise ValueError("AI plan slide_ids and insight_ids must be arrays.")

    if strict:
        if len(raw_slides) > 8 or len(set(raw_slides)) != len(raw_slides):
            raise ValueError("AI plan contains too many or repeated slides.")
        if any(slide_id not in analysis["slides"] for slide_id in raw_slides):
            raise ValueError("AI plan references an unavailable computed slide.")
        if any(insight_id not in analysis["insights"] for insight_id in raw_insights):
            raise ValueError("AI plan references an unavailable computed insight.")

    # Resilient filtering: keep only valid candidates while discarding hallucinated IDs
    valid_slides = [s for s in raw_slides if isinstance(s, str) and s in analysis["slides"]]
    valid_insights = [i for i in raw_insights if isinstance(i, str) and i in analysis["insights"]]

    # Deduplicate while preserving order
    seen = set()
    deduped_slides = []
    for s in valid_slides:
        if s not in seen:
            seen.add(s)
            deduped_slides.append(s)

    required = analysis.get("required_ratings", [])
    ordered = required + [s for s in deduped_slides if s not in required]

    # If the LLM hallucinated entirely or returned nothing valid, fall back to default plan slides
    if not ordered:
        fallback = default_analysis_plan(analysis)
        ordered = fallback["slide_ids"]
    if not valid_insights:
        valid_insights = list(analysis.get("insights", {}).keys())[:4]

    return {"slide_ids": ordered[:8], "insight_ids": valid_insights[:4]}
def _as_number(value):
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        number = float(value)
    else:
        try:
            number = float(str(value).strip().replace(",", "."))
        except (TypeError, ValueError):
            return None
    return number if math.isfinite(number) else None


def _field_title(field, headers, labels):
    label = labels.get(field, field)
    if sum(labels.get(candidate, candidate) == label for candidate in headers) > 1:
        return f"{label} ({field})"
    return label


def _distribution(values, maximum_categories=12):
    counts = Counter(values)
    ordered = counts.most_common(maximum_categories)
    if len(counts) > maximum_categories:
        ordered.append(("Other", sum(count for _, count in counts.most_common()[maximum_categories:])))
    return {str(value): count for value, count in ordered}


def _metric_summary(values, scale):
    count = len(values)
    result = {"n": count, "mean": mean(values), "median": median(values),
              "min": min(values), "max": max(values)}
    if scale == 10:
        top = sum(value >= 9 for value in values)
        bottom = sum(value <= 4 for value in values)
        result.update({"scale": 10, "top_n": top, "top_pct": top / count * 100,
                       "bottom_n": bottom, "bottom_pct": bottom / count * 100,
                       "distribution": {str(score): sum(value == score for value in values)
                                        for score in range(1, 11)}})
    return result


def build_analysis(headers, records, labels):
    total = len(records)
    numeric = {}
    categorical = {}
    for field in headers:
        if not field:
            continue
        values = [row.get(field) for row in records if row.get(field) not in (None, "")]
        if not values:
            continue
        numbers = [_as_number(value) for value in values]
        numeric_values = [value for value in numbers if value is not None]
        name = f"{field} {labels.get(field, field)}".lower()
        identifier = bool(re.search(r"\b(id|identifier|case number|record number)\b", name))
        if (len(numeric_values) == len(values) and not identifier
                and (len(set(numeric_values)) < len(numeric_values) * 0.85 or
                     any(word in name for word in ("rating", "score", "price", "purchase", "intent", "likelihood", "recommend", "evaluate", "age", "income")))):
            unique_values = set(numeric_values)
            scale = 10 if (min(numeric_values) >= 0 and max(numeric_values) <= 10
                           and all(value.is_integer() for value in unique_values)
                           and len(unique_values) <= 11) else None
            numeric[field] = {"field": field, "label": _field_title(field, headers, labels),
                              "values": {index: value for index, value in enumerate(numeric_values)},
                              "summary": _metric_summary(numeric_values, scale), "scale": scale,
                              "answered": len(numeric_values)}
            continue
        counts = Counter(str(value).strip() for value in values)
        if identifier or not 2 <= len(counts) <= 30:
            continue
        categorical[field] = {"field": field, "label": _field_title(field, headers, labels),
                              "counts": dict(counts), "answered": len(values)}

    candidates = {}
    insights = {}
    rating_fields = []
    for field, metric in numeric.items():
        summary = metric["summary"]
        candidate_id = f"metric:{field}"
        metric["id"] = candidate_id
        metric["kind"] = "rating" if metric["scale"] == 10 else "numeric"
        candidates[candidate_id] = metric
        if metric["scale"] == 10:
            rating_fields.append(candidate_id)
            insight_id = f"insight:{field}:overall"
            insights[insight_id] = (
                f"{metric['label']}: mean {summary['mean']:.1f}/10; "
                f"{summary['top_pct']:.0f}% top-box (9-10), "
                f"{summary['bottom_pct']:.0f}% bottom-box (1-4), n={summary['n']}."
            )
            metric["insight_id"] = insight_id

    for field, item in categorical.items():
        candidate_id = f"frequency:{field}"
        groups = _distribution(item["counts"], maximum_categories=8)
        candidates[candidate_id] = {"id": candidate_id, "kind": "frequency", "field": field,
                                    "label": item["label"], "n": item["answered"], "counts": groups}
        item["candidate_id"] = candidate_id

    for metric_field, metric in numeric.items():
        if metric["scale"] != 10:
            continue
        for group_field, group_data in categorical.items():
            groups = {str(value): [] for value in group_data["counts"]}
            for record in records:
                group = record.get(group_field)
                score = _as_number(record.get(metric_field))
                if group is not None and score is not None:
                    groups[str(group).strip()].append(score)
            eligible = {name: values for name, values in groups.items() if len(values) >= 10}
            small_n = sum(len(values) for name, values in groups.items() if len(values) < 10)
            if small_n:
                eligible["Other (<10 per category)"] = [
                    _as_number(record.get(metric_field)) for record in records
                    if record.get(group_field) is not None
                    and len(groups[str(record[group_field]).strip()]) < 10
                    and _as_number(record.get(metric_field)) is not None
                ]
            if len(eligible) < 2 or len(eligible) > 12:
                continue
            candidate_id = f"subgroup:{metric_field}:{group_field}"
            group_stats = {name: _metric_summary(values, 10) for name, values in eligible.items() if values}
            candidates[candidate_id] = {"id": candidate_id, "kind": "subgroup",
                                        "field": metric_field, "label": metric["label"],
                                        "group_field": group_field, "group_label": group_data["label"],
                                        "groups": group_stats}
            best_group = max(group_stats, key=lambda name: group_stats[name]["mean"])
            weakest_group = min(group_stats, key=lambda name: group_stats[name]["mean"])
            insight_id = f"insight:{metric_field}:{group_field}"
            insights[insight_id] = (
                f"{metric['label']} varies by {group_data['label']}: "
                f"{best_group} leads ({group_stats[best_group]['mean']:.1f}/10, n={group_stats[best_group]['n']}); "
                f"{weakest_group} is lowest ({group_stats[weakest_group]['mean']:.1f}/10, n={group_stats[weakest_group]['n']})."
            )
            candidates[candidate_id]["insight_id"] = insight_id

    preferred = [candidate_id for candidate_id in candidates if candidate_id.startswith("metric:")]
    subgroup_ids = [candidate_id for candidate_id in candidates if candidate_id.startswith("subgroup:")]
    frequency_ids = [candidate_id for candidate_id in candidates if candidate_id.startswith("frequency:")]
    # Keep the AI outline compact; retain core rating and a representative subgroup cut.
    ordered_ids = preferred + subgroup_ids[:8] + frequency_ids[:10]
    candidates = {candidate_id: candidates[candidate_id] for candidate_id in ordered_ids[:30]}
    candidate_insights = {candidate["insight_id"]: insights[candidate["insight_id"]]
                          for candidate in candidates.values() if candidate.get("insight_id")}
    return {"n": total, "fields": [labels.get(field, field) for field in headers if field],
            "metrics": numeric, "categories": categorical, "candidates": candidates,
            "insights": candidate_insights, "required": rating_fields}


def default_plan(analysis):
    candidates = analysis["candidates"]
    selected = list(analysis["required"])
    selected.extend(candidate_id for candidate_id, candidate in candidates.items()
                    if candidate["kind"] == "subgroup" and candidate_id not in selected)
    selected.extend(candidate_id for candidate_id, candidate in candidates.items()
                    if candidate["kind"] == "frequency" and candidate_id not in selected)
    selected = selected[:7]
    insight_ids = list(analysis["insights"])[:3]
    return {"slides": selected, "insights": insight_ids}


def local_ai_plan(analysis, settings, model):
    catalog = []
    for candidate_id, candidate in analysis["candidates"].items():
        compact = {key: value for key, value in candidate.items()
                   if key in ("kind", "label", "group_label", "n", "counts", "summary", "groups")}
        catalog.append({"id": candidate_id, **compact})
    instruction = (
        "You are selecting a concise customer research presentation outline. The available records "
        "are only precomputed aggregates, never respondent-level data. Choose relevant slide IDs and "
        "factual insight IDs from the supplied catalog; do not calculate, alter, or invent values. "
        "Always include all required rating slide IDs. Prioritize product ratings and useful buyer/"
        "rejecter cross-tabs, demographic distributions, brand, pricing, and acceptance comparisons "
        "when present. Do not imply "
        "a field exists if it is absent. Return ONLY JSON with keys slides and insights, both arrays "
        "of supplied IDs. Choose up to 7 additional slides and up to 4 insights.\n"
        + json.dumps({"audience": settings.get("Audience", "Research stakeholders"),
                      "respondents": analysis["n"], "fields": analysis["fields"],
                      "required_slides": analysis["required"], "catalog": catalog,
                      "insight_catalog": analysis["insights"]}, ensure_ascii=True)
    )
    body = json.dumps({"model": model, "prompt": instruction, "format": "json", "stream": False,
                       "options": {"temperature": 0}}).encode("utf-8")
    request = Request(f"{OLLAMA_BASE_URL}/api/generate", data=body,
                      headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urlopen(request, timeout=180) as response:
            result = json.load(response)
        return json.loads(result["response"])
    except (URLError, TimeoutError, KeyError, json.JSONDecodeError) as error:
        raise ValueError(f"Local Ollama planning failed: {error}. Start Ollama and check 'ollama list'.") from error


def validate_plan(plan, analysis):
    if not isinstance(plan, dict) or set(plan) != {"slides", "insights"}:
        raise ValueError("Plan must contain only 'slides' and 'insights'.")
    slides, insights = plan["slides"], plan["insights"]
    if (not isinstance(slides, list) or len(slides) > 8 or not isinstance(insights, list)
            or len(insights) > 4):
        raise ValueError("Plan supports at most eight slides and four insights.")
    if any(not isinstance(item, str) or item not in analysis["candidates"] for item in slides):
        raise ValueError("Plan contains a slide that is not in the computed analysis catalog.")
    if len(set(slides)) != len(slides):
        raise ValueError("Plan contains duplicate slide IDs.")
    if not set(analysis["required"]).issubset(slides):
        slides = list(analysis["required"]) + [item for item in slides if item not in analysis["required"]]
    if any(item not in analysis["insights"] for item in insights):
        raise ValueError("Plan contains an insight that is not in the computed facts catalog.")
    return {"slides": slides[:8], "insights": insights}


def new_slide(presentation, title, index, total, source, question, templated,
              backgrounds=None, logo_on_master=False):
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
    text(slide, "AMO  /  CUSTOMER RESEARCH", 0.55, 0.30, 6.5, 0.30, 10, TEAL, True,
         font_name=HEADING_FONT)
    text(slide, title[:90], 0.55, 0.87, 12.1, 0.88, 28, TEAL, True,
         font_name=HEADING_FONT)
    text(slide, f"Source: {source}"[:150], 0.55, 6.48, 9.9, 0.20, 8, MUTED)
    text(slide, f"Question answered: {question}"[:220], 0.55, 6.70, 9.9, 0.34, 8, MUTED)
    text(slide, str(index), 0.37, 7.14, 0.35, 0.20, 9, MUTED)
    if BRAND_LOGO.is_file() and not logo_on_master:
        slide.shapes.add_picture(str(BRAND_LOGO), Inches(11.16), Inches(6.74),
                                 width=Inches(1.8), height=Inches(0.62))
    return slide


def add_takeaway(slide, message):
    background = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.68), Inches(1.78),
                                        Inches(11.95), Inches(0.48))
    background.fill.solid()
    background.fill.fore_color.rgb = LIGHT_BLUE
    background.line.fill.background()
    accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.68), Inches(1.78),
                                    Inches(0.06), Inches(0.48))
    accent.fill.solid()
    accent.fill.fore_color.rgb = TEAL
    accent.line.fill.background()
    text(slide, "TAKEAWAY", 0.88, 1.84, 1.15, 0.18, 9, TEAL, True)
    text(slide, message[:185], 2.0, 1.82, 10.35, 0.34, 12, INK, True)


def add_editable_bar_chart(slide, categories, values, x, y, width, height, maximum=None,
                          color=TEAL, number_format="0"):
    chart_data = CategoryChartData()
    chart_data.categories = categories
    chart_data.add_series("Responses", values)
    chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y),
                                   Inches(width), Inches(height), chart_data).chart
    chart.has_legend = False
    chart.has_title = False
    chart.chart_style = 10
    chart.value_axis.minimum_scale = 0
    chart.value_axis.maximum_scale = maximum or max(values, default=1) * 1.15
    chart.plots[0].series[0].format.fill.solid()
    chart.plots[0].series[0].format.fill.fore_color.rgb = color
    chart.plots[0].has_data_labels = True
    chart.plots[0].data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    chart.plots[0].data_labels.number_format = number_format
    chart.plots[0].data_labels.number_format_is_linked = False
    chart.value_axis.tick_labels.font.name = BODY_FONT
    chart.category_axis.tick_labels.font.name = BODY_FONT
    return chart


def add_editable_stacked_chart(slide, item, x, y, width, height):
    chart_data = CategoryChartData()
    chart_data.categories = [group["name"] for group in item["groups"]]
    for outcome in item["outcomes"]:
        chart_data.add_series(outcome, [group["counts"].get(outcome, 0) for group in item["groups"]])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, Inches(x), Inches(y),
                                   Inches(width), Inches(height), chart_data).chart
    chart.has_legend = True
    chart.legend.include_in_layout = False
    chart.chart_style = 10
    chart.value_axis.minimum_scale = 0
    chart.value_axis.tick_labels.font.name = BODY_FONT
    chart.category_axis.tick_labels.font.name = BODY_FONT
    chart.plots[0].has_data_labels = True
    chart.plots[0].data_labels.position = XL_LABEL_POSITION.CENTER
    return chart


def add_rating_gauge(slide, rating, x, y, size):
    value = rating["summary"]["mean"]
    chart_data = CategoryChartData()
    chart_data.categories = ["Average rating", "Remaining", "Hidden half"]
    chart_data.add_series("Rating", [value, max(0, 10 - value), 10])
    chart = slide.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(x), Inches(y),
                                   Inches(size), Inches(size), chart_data).chart
    chart.has_legend = False
    chart.has_title = False
    chart.chart_style = 10
    plot_element = chart.plots[0]._element
    plot_element.find(qn("c:firstSliceAng")).set("val", "270")
    plot_element.find(qn("c:holeSize")).set("val", "72")
    palette = (TEAL, LIGHT_BLUE, PALE)
    for point, color in zip(chart.series[0].points, palette):
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = color
        point.format.line.fill.background()
    text(slide, f"{value:.1f}", x + size * 0.27, y + size * 0.39,
            size * 0.46, 0.42, 25, TEAL, True, font_name=HEADING_FONT)
    text(slide, "Average rating", x + size * 0.16, y + size * 0.57,
         size * 0.68, 0.24, 10, INK, True)
    text(slide, "Score on 10-point scale", x + size * 0.06, y + size * 0.70,
         size * 0.88, 0.22, 8, MUTED)
    return chart


def add_metric(slide, label, value, x, width=3.7, y=2.39):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                   Inches(width), Inches(0.96))
    shape.fill.solid()
    shape.fill.fore_color.rgb = PALE
    shape.line.color.rgb = LIGHT_BLUE
    text(slide, value, x + 0.16, y + 0.10, width - 0.30, 0.40, 25, TEAL, True,
         font_name=HEADING_FONT)
    text(slide, label, x + 0.16, y + 0.56, width - 0.30, 0.25, 10, MUTED)


def add_group_stats_table(slide, groups, top=5.0, x=0.7, width=11.9, compact=False):
    headers = ["Subgroup", "Base (n)", "Mean / 10", "Top-box 9-10", "Bottom-box 1-4"]
    table = slide.shapes.add_table(len(groups) + 1, len(headers), Inches(x), Inches(top),
                                  Inches(width), Inches(min(1.72, 0.30 * (len(groups) + 1)))).table
    base_widths = (3.5, 1.2, 1.5, 2.25, 2.45) if compact else (3.8, 1.35, 1.7, 2.35, 2.7)
    scale = width / sum(base_widths)
    widths = tuple(value * scale for value in base_widths)
    for column, (header, width) in enumerate(zip(headers, widths)):
        table.columns[column].width = Inches(width)
        cell = table.cell(0, column)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = TEAL
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.name = BODY_FONT
            paragraph.font.size = Pt(8 if compact else 10)
            paragraph.font.bold = True
            paragraph.font.color.rgb = RGBColor(255, 255, 255)
    for row_index, group in enumerate(groups, start=1):
        values = [group["name"], str(group["n"]), f"{group['mean']:.1f}",
                  f"{group['top_share']:.0%} ({group['top_count']})",
                  f"{group['bottom_share']:.0%} ({group['bottom_count']})"]
        for column, value in enumerate(values):
            cell = table.cell(row_index, column)
            cell.text = value
            cell.fill.solid()
            cell.fill.fore_color.rgb = PALE if row_index % 2 else RGBColor(255, 255, 255)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.name = BODY_FONT
                paragraph.font.size = Pt(8 if compact else 10)
                paragraph.font.color.rgb = INK


def _master_has_amo_logo(presentation):
    if not BRAND_LOGO.is_file():
        return False
    logo_bytes = BRAND_LOGO.read_bytes()
    for master in presentation.slide_masters:
        for shape in master.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE and shape.image.blob == logo_bytes:
                return True
    return False


def _coverage_note(coverage):
    unavailable = []
    for key, label in (("demographics", "demographics"),
                       ("purchase_intent_or_rejection", "purchase intent/rejection"),
                       ("pricing_expectations", "pricing expectations"),
                       ("emerging_brand_acceptance", "emerging-brand acceptance")):
        if not coverage[key]:
            unavailable.append(label)
    if unavailable:
        return "Not captured in this file: " + ", ".join(unavailable) + "."
    return "Buyer, rejection, pricing and brand-acceptance measures are present in this file."


def _rating_takeaway(rating, analysis, plan):
    summary = rating["summary"]
    top_share = summary["top_share"]
    if top_share >= 0.5:
        base = f"The design resonates with a majority: {top_share:.0%} rate it 9-10."
    else:
        base = f"{top_share:.0%} rate the design 9-10; the mean is {summary['mean']:.1f}/10."
    cuts = [analysis["slides"][slide_id] for slide_id in plan["slide_ids"]
            if slide_id in analysis["slides"]
            and analysis["slides"][slide_id]["kind"] == "subgroup"
            and analysis["slides"][slide_id]["field"] == rating["field"]]
    drive_cuts = [cut for cut in cuts if "drive" in cut["group_field"].lower()
                  or "engine" in cut["group_field"].lower()
                  or "drive" in cut["group_label"].lower()]
    candidate_cuts = drive_cuts or cuts
    if candidate_cuts:
        cut = candidate_cuts[0]
        leader = max(cut["groups"], key=lambda group: group["top_share"])
        if leader["top_share"] > top_share and leader["n"] >= 10:
            segment = "current BEV owners" if leader["name"] == "BEV" else leader["name"]
            base = base[:-1] + f", rising to {leader['top_share']:.0%} among {segment}."
    return base


def render_analysis(settings, analysis, plan, output, template=None):
    presentation = Presentation(template) if template else Presentation()
    if template and abs(presentation.slide_width / presentation.slide_height - 16 / 9) > 0.02:
        raise ValueError("Template must be a 16:9 PowerPoint presentation.")
    backgrounds = []
    if template:
        for example in presentation.slides:
            try:
                color = example.background.fill.fore_color.rgb
                if color:
                    backgrounds.append(color)
            except (AttributeError, TypeError):
                pass
        for slide_id in list(presentation.slides._sldIdLst):
            presentation.part.drop_rel(slide_id.rId)
            presentation.slides._sldIdLst.remove(slide_id)
        logo_on_master = _master_has_amo_logo(presentation)
    else:
        presentation.slide_width = Inches(13.333)
        presentation.slide_height = Inches(7.5)
        logo_on_master = False

    rating_ids = [item for item in plan["slide_ids"] if item in analysis["required_ratings"]]
    detail_ids = [item for item in plan["slide_ids"] if item not in analysis["required_ratings"]]
    total_slides = len(rating_ids) + len(detail_ids)
    source = settings.get("Source name", settings.get("Data sheet", "Data"))
    index = 1
    for rating_id in rating_ids:
        rating = analysis["slides"][rating_id]
        summary = rating["summary"]
        headline = (f"{rating['display_label']} earns strong ratings"
                    if summary["mean"] >= 7 else f"{rating['display_label']} rating overview")
        slide = new_slide(presentation, headline, index, total_slides,
                          source, rating["label"], bool(template), backgrounds, logo_on_master)
        add_takeaway(slide, _rating_takeaway(rating, analysis, plan))
        add_rating_gauge(slide, rating, 0.75, 2.24, 3.35)

        subgroup_cut = next((analysis["slides"][slide_id] for slide_id in plan["slide_ids"]
                             if slide_id in analysis["slides"]
                             and analysis["slides"][slide_id]["kind"] == "subgroup"
                             and analysis["slides"][slide_id]["field"] == rating["field"]
                             and ("drive" in analysis["slides"][slide_id]["group_field"].lower()
                                  or "engine" in analysis["slides"][slide_id]["group_field"].lower())), None)
        add_metric(slide, "Respondent base", str(analysis["n"]), 4.48, 2.45, y=2.42)
        add_metric(slide, "Top-box · 9-10", f"{summary['top_share']:.0%}", 7.13, 2.45, y=2.42)
        add_metric(slide, "Bottom-box · 1-4", f"{summary['bottom_share']:.0%}", 9.78, 2.65, y=2.42)
        if subgroup_cut:
            text(slide, f"BY {subgroup_cut['group_label'].upper()}", 4.52, 3.72, 7.5, 0.24,
                 10, TEAL, True)
            add_group_stats_table(slide, subgroup_cut["groups"], top=4.02, x=4.42,
                                  width=8.18, compact=True)
        text(slide, _coverage_note(analysis["coverage"]), 0.72, 6.27, 11.8, 0.34, 9, MUTED)
        index += 1

    for candidate_id in detail_ids:
        item = analysis["slides"][candidate_id]
        if item["kind"] == "subgroup":
            title = f"{_display_question(item['label'])} rating by {item['group_label'].lower()}"
            question = item["question"]
            groups = sorted(item["groups"], key=lambda row: row["mean"], reverse=True)
            best, lowest = groups[0], groups[-1]
            takeaway = (f"{best['name']} scores highest ({best['mean']:.1f}/10, n={best['n']}); "
                        f"{lowest['name']} is lowest ({lowest['mean']:.1f}/10, n={lowest['n']}).")
        elif item["kind"] == "frequency":
            title = f"Respondent profile: {item['label']}"
            question = item.get("question", f"How is the respondent base distributed across {item['label']}?")
            groups = None
            entries = sorted(item["counts"].items(), key=lambda pair: pair[1], reverse=True)
            leading, count = entries[0]
            takeaway = f"Most common: {leading}, {count} of {item['n']} respondents ({count / item['n']:.0%})."
        elif item["kind"] == "numeric":
            title = f"Respondent profile: {item['label']}"
            question = f"How are respondents distributed across {item['label']}?"
            groups = None
            entries = list(item["counts"].items())
            takeaway = f"Average {item['mean']:.1f}; median {item['median']:.1f}; n={item['n']}."
        else:
            title = f"{item['label']} by {item['group_label']}"
            question = f"How does {item['label']} vary across {item['group_label']}?"
            groups = None
            largest = max(item["groups"], key=lambda group: group["n"])
            takeaway = f"Largest segment: {largest['name']} ({largest['n']} responses)."

        slide = new_slide(presentation, title, index, total_slides, source, question,
                              bool(template), backgrounds, logo_on_master)
        add_takeaway(slide, takeaway)
        if item["kind"] == "subgroup":
            text(slide, f"{item['label']} · mean rating / 10", 0.65, 2.40, 11.8, 0.25, 11, MUTED, True)
            names = [f"{group['name']} (n={group['n']})" for group in groups]
            add_editable_bar_chart(slide, names, [group["mean"] for group in groups],
                                   0.70, 2.72, 11.9, 2.15, maximum=10.5,
                                   color=CORAL, number_format="0.0")
            add_group_stats_table(slide, groups, top=5.00)
        elif item["kind"] in ("frequency", "numeric"):
            text(slide, f"{item['n']} answered · case frequency by category", 0.65, 2.40,
                 11.8, 0.25, 11, MUTED, True)
            add_editable_bar_chart(slide, [name for name, _ in entries], [count for _, count in entries],
                                   0.70, 2.72, 11.9, 3.85,
                                   maximum=max(value for _, value in entries) * 1.2,
                                   color=CORAL if item["kind"] == "numeric" else TEAL)
        else:
            text(slide, "Response frequencies by customer segment", 0.65, 2.40,
                 11.8, 0.25, 11, MUTED, True)
            add_editable_stacked_chart(slide, item, 0.70, 2.72, 11.9, 3.85)
        index += 1
    presentation.save(output)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", type=Path, default=Path("examples/data-set.csv"))
    parser.add_argument("--output", type=Path, help="Override the workbook output name")
    parser.add_argument("--ai", action="store_true", help="Use local Ollama to select slides")
    parser.add_argument("--model", default="llama3.1:8b", help="Installed Ollama model name")
    parser.add_argument("--plan-only", action="store_true", help="Print the validated plan without writing slides")
    args = parser.parse_args()
    try:
        settings, headers, records, labels = load_survey(args.input)
        analysis = analyze_survey(headers, records, labels)
        plan = (ai_analysis_plan(analysis, settings, args.model) if args.ai
                else default_analysis_plan(analysis))
        plan = validate_analysis_plan(plan, analysis)
        if args.plan_only:
            print(json.dumps(plan, indent=2))
            return
        name = settings["Output name"]
        if args.input.suffix.lower() == ".csv":
            name = f"{args.input.stem}-presentation.pptx"
        if not name or Path(name).name != name or Path(name).suffix.lower() != ".pptx":
            raise ValueError("Output name must be a .pptx filename with no directory.")
        output = args.output or args.input.parent / name
        if output.resolve() == args.input.resolve():
            raise ValueError("Output presentation cannot replace the input workbook.")
        template = args.input.parent / settings["Template"] if settings["Template"] else None
        render_analysis(settings, analysis, plan, output, template)
        print(f"Created {output} with {len(plan['slide_ids'])} editable slides from {analysis['n']} responses")
    except (ValueError, FileNotFoundError) as error:
        parser.error(str(error))


if __name__ == "__main__":
    main()