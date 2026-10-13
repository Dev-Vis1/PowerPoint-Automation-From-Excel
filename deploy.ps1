param(
    [string]$Model = "llama3.1:8b",
    [switch]$SkipModelPull
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    throw "Docker CLI was not found. Install and start Docker Desktop, then reopen PowerShell."
}

docker info *> $null
if ($LASTEXITCODE -ne 0) {
    throw "Docker Engine is unavailable. Start Docker Desktop and wait until its engine is running."
}

docker compose up -d --build
if ($LASTEXITCODE -ne 0) {
    throw "Docker Compose failed to build or start the services. Check the output above."
}

if (-not $SkipModelPull) {
    Write-Host "Waiting for the Ollama service..."
    $ready = $false
    for ($attempt = 0; $attempt -lt 60; $attempt++) {
        docker compose exec -T ollama ollama list *> $null
        if ($LASTEXITCODE -eq 0) {
            $ready = $true
            break
        }
        Start-Sleep -Seconds 2
    }

    if (-not $ready) {
        throw "Ollama did not become ready within 120 seconds. Check: docker compose logs ollama"
    }

    Write-Host "Ensuring Ollama model '$Model' is available..."
    docker compose exec -T ollama ollama pull $Model
    if ($LASTEXITCODE -ne 0) {
        throw "Could not pull model '$Model'. The app remains available with AI planning disabled or its deterministic fallback."
    }
}

Write-Host "Survey Studio is available at http://127.0.0.1:8001"
if ($SkipModelPull) {
    Write-Host "Model download skipped. Pull a model later with: docker compose exec ollama ollama pull $Model"
}
