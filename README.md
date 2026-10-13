# Excel survey to editable PowerPoint

`survey_studio.smart_slides` analyzes survey files and builds editable PowerPoint decks. Python calculates all frequencies, means, score shares, and subgroup statistics. Optional local Ollama AI selects relevant computed subgroup slides and their order; it never supplies or changes figures. No subscription or cloud API key is required.

## Layout

- `survey_studio/`: web app, slide generation, and static assets
- `tests/`: generator and web workflow tests
- `examples/`: sample workbooks and presentations
- `storage/`: temporary uploads and seven-day downloads (ignored by Git)

## Docker Compose (App + Ollama Stack)

You can run the entire stack (FastAPI web app + containerized Ollama) with Docker Compose:

```bash
docker compose up -d --build
```

On first startup, pull the Ollama model into the container's volume:

```bash
docker exec -it survey-ollama ollama pull llama3.1:8b
```

Access the app at `http://127.0.0.1:8001`. The containerized app automatically routes AI requests to the `ollama` container at `http://ollama:11434`. Generated decks and downloaded models persist in Docker volumes (`survey_storage` and `ollama_models`).

## Web app

Install dependencies and start the local upload workspace:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install openpyxl python-pptx fastapi uvicorn python-multipart httpx
.\.venv\Scripts\python.exe -m uvicorn survey_studio.app:app --host 127.0.0.1 --port 8000
```

Open `http://127.0.0.1:8000/`. Upload an `.xlsx` workbook or `.csv` survey file (up to 15 MB), select its worksheet (`Data` for CSV), and generate the presentation. CSV supports one header row or two header rows: field IDs followed by question wording. Choose **Upload template** for a 16:9 `.pptx` or select an available design. Example slides are removed; the theme/master and solid backgrounds are retained. Local Ollama planning is on by default and chooses among Python-computed subgroups. Turn it off for deterministic slide selection.

Uploads are held under `storage/uploads` only during inspection, preview, or generation, then deleted. Finished decks can be downloaded for seven days. Cleanup runs at startup and hourly. This is a **local, single-user** workflow, not an authenticated public hosting setup. Uploaded records and comments are not sent to Ollama; the local planner receives question labels and computed aggregates.

## Command line

Run these commands from the project root with the dependencies installed above. The default input is the supplied real file `examples/data-set.csv`; set `--input` to use another workbook or CSV.

For local AI, install Ollama, download a model once, and generate slides:

```powershell
ollama pull llama3.1:8b
.\.venv\Scripts\python.exe -m survey_studio.smart_slides --input "examples/data-set.csv" --output "examples/Exterior design analysis.pptx" --ai
```

Ollama must be running locally. Model download and inference use your disk and computer, not a subscription. If you already have another Ollama model, pass `--model your-model-name`. Inspect the selected slides before generating them with `--ai --plan-only`. Without `--ai`, Ollama is not needed:

```powershell
.\.venv\Scripts\python.exe -m survey_studio.smart_slides --input "examples/data-set.csv" --plan-only
.\.venv\Scripts\python.exe -m survey_studio.smart_slides --input "examples/data-set.csv" --output "examples/Exterior design analysis.pptx"
```

## Workbook layout

CSV can have one header row or two header rows (field IDs followed by question wording); comma and semicolon delimiters are detected. Numeric integer 0-10/1-10 rating questions generate the mean, score distribution, top-box (9-10), and bottom-box (1-4) shares. Categorical questions generate response-frequency charts and rating subgroup cuts. Subgroups with fewer than 10 responses are combined into `Other` to avoid unstable tiny bases. Charts and subgroup tables are native editable PowerPoint objects.

The supplied automotive dataset contains 203 cases, current brand/model, drive type, and exterior-design ratings. It does not contain demographics, pricing expectations, purchase intent/rejection, or emerging-brand acceptance. The deck calls out those gaps rather than inventing findings. Add those questions to future research data to enable those analyses.

The optional `Presentation` sheet has `Setting` in A1 and `Value` in B1. Supported rows:

| Setting | Example | Purpose |
| --- | --- | --- |
| Output name | `quarterly_results.pptx` | Generated file beside the workbook |
| Title | `Employee survey results` | Overview title |
| First slide | `Executive summary` | First-slide heading |
| Audience | `Team leads` | Overview audience and AI planning context |
| Data sheet | `Responses` | Worksheet containing answers |
| Template | `company-template.pptx` | Optional 16:9 PowerPoint design reference beside the workbook |

Templates must be 16:9. Example slides are removed; the template theme/master and solid backgrounds are retained. To override the output path, use `--output other-report.pptx`. The deck is **not linked to its source data**. Rerunning replaces a same-named output, so use another name to preserve edits.

Run all tests with `.\.venv\Scripts\python.exe -m unittest discover -s tests -v`.