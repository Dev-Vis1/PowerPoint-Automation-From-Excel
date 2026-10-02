# Excel survey to editable PowerPoint

`survey_studio.smart_slides` inspects a survey workbook and builds an editable deck from its columns. It uses Python to count responses and create native PowerPoint charts and text. Optional local Ollama AI chooses which questions deserve slides, their order, and chart orientation. No subscription or cloud API key is required. Without AI, it uses a deterministic plan. `survey_studio.survey_to_slides` is the original fixed, three-slide example.

## Layout

- `survey_studio/`: web app, slide generation, and static assets
- `tests/`: generator and web workflow tests
- `examples/`: sample workbooks and presentations
- `storage/`: temporary uploads and seven-day downloads (ignored by Git)

## Web app

Install dependencies and start the local upload workspace:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install openpyxl python-pptx fastapi uvicorn python-multipart httpx
.\.venv\Scripts\python.exe -m uvicorn survey_studio.app:app --host 127.0.0.1 --port 8000
```

Open `http://127.0.0.1:8000/`. Choose an `.xlsx` workbook (up to 15 MB), select its data worksheet, and generate a PowerPoint. Choose **Upload template** for your own 16:9 `.pptx`, or click a name under **Available templates** to select and preview it. Use the arrows to page through multi-slide template previews. The preview represents text and solid colors/shapes, not a pixel-perfect PowerPoint rendering. Example slides in an uploaded template are removed from the output; the generated report retains its theme/master and cycles through solid slide backgrounds. Shapes on example slides are not copied to output slides, so put persistent branding on the template's slide master. Local AI planning is on by default; it uses the installed `llama3.1:8b` Ollama model. Start Ollama and run `ollama pull llama3.1:8b` once if needed. Turn the switch off for deterministic slide selection with no Ollama dependency. The web app honors the `Presentation` sheet's title, first slide, audience, and output filename, while the selected worksheet overrides its `Data sheet` value.

Uploads (workbooks and templates) are held under `storage/uploads` only during inspection, preview, or generation and deleted when each request finishes, including failed requests. Finished decks are stored under `storage/generated`, can be downloaded by their unguessable link, and are deleted after seven days. Cleanup runs at startup and hourly while the app is running; expired files are also rejected and removed when requested. This is a **local, single-user** workflow, not an authenticated public hosting setup. Do not expose the server to the internet with sensitive survey data. Uploaded workbooks are not sent to Ollama; the local planner receives question labels and aggregate counts.

## Command line

Run these commands from the project root with the dependencies installed above:

The supplied `examples/demo_survey.xlsx` has 24 fictional responses plus a `Presentation` control sheet. `examples/survey_responses.xlsx` is also supported, even though it predates the optional control sheet. To generate another sample without overwriting the supplied files:

```powershell
.\.venv\Scripts\python.exe -m survey_studio.smart_slides --sample --input examples/new_survey.xlsx --output examples/new_survey.pptx
```

For local AI, install Ollama, download a model once, and generate slides:

```powershell
ollama pull llama3.1:8b
.\.venv\Scripts\python.exe -m survey_studio.smart_slides --input examples/demo_survey.xlsx --output examples/ai_report.pptx --ai
```

Ollama must be running locally. Model download and inference use your disk and computer, not a subscription. If you already have another Ollama model, pass `--model your-model-name`. Inspect the selected slides before generating them with `--ai --plan-only`. Without `--ai`, Ollama is not needed:

```powershell
.\.venv\Scripts\python.exe -m survey_studio.smart_slides --input examples/demo_survey.xlsx --plan-only
.\.venv\Scripts\python.exe -m survey_studio.smart_slides --input examples/demo_survey.xlsx --output examples/new_report.pptx
```

## Workbook layout

Put question names in the first row of the data sheet and one response per subsequent row. Headers need not match the sample. The generator detects categorical questions (2-12 answers, with sufficiently repeated values) and free-text comments; continuous numerical columns and unique identifiers are skipped. Missing answers are allowed. Up to four chart slides plus one comment slide are selected. Python calculates every displayed count; AI receives only question labels and aggregate counts, not individual responses or comments.

The optional `Presentation` sheet has `Setting` in A1 and `Value` in B1. Supported rows:

| Setting | Example | Purpose |
| --- | --- | --- |
| Output name | `quarterly_results.pptx` | Generated file beside the workbook |
| Title | `Employee survey results` | Overview title |
| First slide | `Executive summary` | First-slide heading |
| Audience | `Team leads` | Overview audience and AI planning context |
| Data sheet | `Responses` | Worksheet containing answers |
| Template | `company-template.pptx` | Optional empty 16:9 PowerPoint template beside the workbook |

The template must contain no slides; its master/background are retained, while the generator draws its own text, accent colors, and charts on a blank layout. To override the workbook output location, use `--output other-report.pptx`. The generated deck is **not linked to Excel**. Rerunning replaces the output presentation, so choose a new output name if you have edited the previous deck in PowerPoint. Review selected charts and comments for accuracy and privacy before sharing.

Run all tests with `.\.venv\Scripts\python.exe -m unittest discover -s tests -v`.