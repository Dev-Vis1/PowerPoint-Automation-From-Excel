# Excel survey to editable PowerPoint

`survey_studio.smart_slides` analyzes survey files and builds editable PowerPoint decks. Python calculates all frequencies, means, score shares, and subgroup statistics. Optional local Ollama AI selects relevant computed subgroup slides and their order; it never supplies or changes figures. No subscription or cloud API key is required.

## Layout

- `survey_studio/`: web app, slide generation, and static assets
- `tests/`: generator and web workflow tests
- `examples/`: sample workbooks and presentations
- `storage/`: temporary uploads and seven-day downloads (ignored by Git)

## Docker Deployment

The Docker deployment includes the FastAPI app and a separate Ollama service. Python packages are installed in the app image; Ollama model files are kept in a persistent Docker volume, not bundled into the image. The first model download is several gigabytes and requires internet access.

### Windows quick start

Start Docker Desktop and wait until the Docker engine is running. From the project directory, run:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\deploy.ps1
```

The script builds and starts the app and Ollama containers, waits for Ollama, pulls `llama3.1:8b` if needed, and prints the app URL: `http://127.0.0.1:8001`.

To use a smaller model or skip downloading a model during deployment:

```powershell
.\deploy.ps1 -Model llama3.2:3b
.\deploy.ps1 -SkipModelPull
```

When model pulling is skipped, download one later with `docker compose exec ollama ollama pull llama3.1:8b`. The app can still generate decks without Ollama planning; Python computes the statistics and the deterministic planner selects slides.

### Docker Compose commands

The deployment can also be managed directly:

```powershell
docker compose up -d --build
docker compose exec ollama ollama pull llama3.1:8b
docker compose ps
docker compose logs -f survey-studio ollama
docker compose down
```

`docker compose down` stops and removes the containers but keeps the named `survey_storage` and `ollama_models` volumes. To remove those volumes as well, run `docker compose down --volumes`; this permanently deletes saved decks and downloaded models.

Compose routes the app to Ollama through `http://ollama:11434`. You can change the model with the `OLLAMA_MODEL` setting in `docker-compose.yml`. The Compose file includes an optional NVIDIA GPU reservation, commented out by default; enable it only on a Docker host configured for NVIDIA Container Toolkit. Without GPU access, Ollama runs on CPU and may respond more slowly.

If Docker reports a 500 error for `dockerDesktopLinuxEngine/_ping`, that is a Docker Desktop engine startup issue, not an application build error. Start or restart Docker Desktop, then check `docker version` and confirm it prints both **Client** and **Server** sections before running deployment again.

## Web app

For a non-containerized local run, install dependencies and start the upload workspace:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install openpyxl python-pptx fastapi uvicorn python-multipart httpx
.\.venv\Scripts\python.exe -m uvicorn survey_studio.app:app --host 127.0.0.1 --port 8000
```

Open `http://127.0.0.1:8000/`. Upload an `.xlsx` workbook or `.csv` survey file (up to 15 MB), select its worksheet (`Data` for CSV), and generate the presentation. CSV supports one header row or two header rows: field IDs followed by question wording. Choose **Upload template** for a 16:9 `.pptx` or select an available design. Example slides are removed; the theme/master and solid backgrounds are retained. Local Ollama planning is on by default and chooses among Python-computed slide candidates. Turn it off for deterministic slide selection. For local Ollama, install Ollama on the host, pull a model, and set `OLLAMA_BASE_URL` if it is not using the default `http://127.0.0.1:11434`.

Uploads are held under `storage/uploads` only during inspection, preview, or generation, then deleted. Finished decks can be downloaded for seven days. Cleanup runs at startup and hourly. This is a **local, single-user** workflow, not an authenticated public hosting setup. Uploaded respondent rows are not sent to Ollama; the planner receives computed candidate summaries and identifiers.

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