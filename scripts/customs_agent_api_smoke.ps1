param(
    [ValidateSet("create", "claim", "upload", "status", "result")]
    [string]$Action = "create",
    [string]$BaseUrl = "http://127.0.0.1:8765",
    [string]$TaskId = "",
    [string[]]$Pdf = @(),
    [string]$AgentName = "finance-pc-01",
    [string]$Flow = "both",
    [string]$StartDate = "",
    [string]$EndDate = "",
    [string]$IdempotencyKey = ""
)

$ErrorActionPreference = "Stop"
$projectRoot = Split-Path -Parent $PSScriptRoot
$envFile = Join-Path $projectRoot ".env"

function Read-DotEnv {
    $values = @{}
    if (-not (Test-Path -LiteralPath $envFile)) {
        return $values
    }
    foreach ($raw in Get-Content -LiteralPath $envFile -Encoding UTF8) {
        $line = $raw.Trim()
        if (-not $line -or $line.StartsWith("#") -or -not $line.Contains("=")) {
            continue
        }
        $name, $value = $line.Split("=", 2)
        $values[$name.Trim()] = $value.Trim()
    }
    return $values
}

function Get-Setting([string]$Name) {
    $value = [Environment]::GetEnvironmentVariable($Name)
    if ($value) {
        return $value
    }
    return [string]$script:dotenv[$Name]
}

function Show-Json($Value) {
    $Value | ConvertTo-Json -Depth 20
}

$script:dotenv = Read-DotEnv
$base = $BaseUrl.TrimEnd("/")
$createToken = Get-Setting "CUSTOMS_TASK_CREATE_TOKEN"
$agentToken = Get-Setting "CUSTOMS_AGENT_TOKEN"

switch ($Action) {
    "create" {
        if (-not $createToken) {
            throw "CUSTOMS_TASK_CREATE_TOKEN is not configured"
        }
        if (-not $IdempotencyKey) {
            $IdempotencyKey = [guid]::NewGuid().ToString()
        }
        $body = @{
            authToken = $createToken
            source = "manual_smoke_test"
            workflow = "customs"
            flow = $Flow
            startDate = $StartDate
            endDate = $EndDate
            idempotencyKey = $IdempotencyKey
        }
        $response = Invoke-RestMethod `
            -Method Post `
            -Uri "$base/api/customs/tasks" `
            -ContentType "application/json; charset=utf-8" `
            -Body ($body | ConvertTo-Json)
        Show-Json $response
    }

    "claim" {
        if (-not $agentToken) {
            throw "CUSTOMS_AGENT_TOKEN is not configured"
        }
        $response = Invoke-RestMethod `
            -Method Get `
            -Uri "$base/api/customs/tasks/claim" `
            -Headers @{
                Authorization = "Bearer $agentToken"
                "X-Agent-Name" = $AgentName
            }
        Show-Json $response
    }

    "upload" {
        if (-not $agentToken) {
            throw "CUSTOMS_AGENT_TOKEN is not configured"
        }
        if (-not $TaskId) {
            throw "upload requires -TaskId"
        }
        if ($Pdf.Count -eq 0) {
            throw "upload requires at least one -Pdf file"
        }
        foreach ($path in $Pdf) {
            if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
                throw "PDF not found: $path"
            }
            if ([IO.Path]::GetExtension($path).ToLowerInvariant() -ne ".pdf") {
                throw "Not a PDF: $path"
            }
        }

        $curlArgs = @(
            "--silent",
            "--show-error",
            "--fail-with-body",
            "-X", "POST",
            "$base/api/customs/tasks/$TaskId/upload",
            "-H", "Authorization: Bearer $agentToken",
            "-H", "X-Agent-Name: $AgentName",
            "-F", "agent_name=$AgentName"
        )
        foreach ($path in $Pdf) {
            $curlArgs += @("-F", "files=@$path;type=application/pdf")
        }
        $raw = & curl.exe @curlArgs
        if ($LASTEXITCODE -ne 0) {
            throw "Upload failed: $raw"
        }
        Show-Json ($raw | ConvertFrom-Json)
    }

    "status" {
        if (-not $agentToken) {
            throw "CUSTOMS_AGENT_TOKEN is not configured"
        }
        if (-not $TaskId) {
            throw "status requires -TaskId"
        }
        $response = Invoke-RestMethod `
            -Method Get `
            -Uri "$base/api/customs/tasks/$TaskId" `
            -Headers @{ Authorization = "Bearer $agentToken" }
        Show-Json $response
    }

    "result" {
        if (-not $agentToken) {
            throw "CUSTOMS_AGENT_TOKEN is not configured"
        }
        if (-not $TaskId) {
            throw "result requires -TaskId"
        }
        $response = Invoke-RestMethod `
            -Method Get `
            -Uri "$base/api/customs/tasks/$TaskId/result" `
            -Headers @{ Authorization = "Bearer $agentToken" }
        Show-Json $response
    }
}
