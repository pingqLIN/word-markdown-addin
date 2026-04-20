param(
  [Parameter(Mandatory = $true, Position = 0)]
  [string]$MarkdownPath,
  [string]$WordPath = "${env:ProgramFiles}\Microsoft Office\root\Office16\WINWORD.EXE"
)

$resolvedMarkdownPath = Resolve-Path -LiteralPath $MarkdownPath -ErrorAction Stop

if (-not (Test-Path $WordPath)) {
  throw "Word executable not found: $WordPath"
}

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
$pendingDirectory = Join-Path $repoRoot.Path ".local"
$pendingPath = Join-Path $pendingDirectory "pending-open.json"
$devServerUrl = "http://localhost:3000/taskpane.html"
$compatibilityUrl = "http://localhost:3000/api/pending-markdown"

function Test-TaskpaneEndpoint {
  try {
    $response = Invoke-WebRequest -Uri $devServerUrl -UseBasicParsing -TimeoutSec 2
    if ($response.StatusCode -ne 200) {
      return $false
    }

    $compatibilityResponse = Invoke-WebRequest -Uri $compatibilityUrl -UseBasicParsing -TimeoutSec 2 -SkipHttpErrorCheck
    return $compatibilityResponse.StatusCode -eq 200 -or $compatibilityResponse.StatusCode -eq 204
  } catch {
    return $false
  }
}

function Start-DevServerIfNeeded {
  if (Test-TaskpaneEndpoint) {
    return
  }

  $node = Get-Command node -ErrorAction SilentlyContinue
  if (-not $node) {
    Write-Warning "node was not found. The local dev server could not be started automatically. Run npm run setup first."
    return
  }

  Start-Process -FilePath $node.Source -ArgumentList "scripts/dev-server.js" -WorkingDirectory $repoRoot.Path -WindowStyle Hidden | Out-Null
  Start-Sleep -Seconds 2

  if (-not (Test-TaskpaneEndpoint)) {
    Write-Warning "The dev server was started, but http://localhost:3000/taskpane.html is still not reachable."
  }
}

New-Item -ItemType Directory -Path $pendingDirectory -Force | Out-Null

$payload = [ordered]@{
  fileName = [System.IO.Path]::GetFileName($resolvedMarkdownPath.Path)
  fullPath = $resolvedMarkdownPath.Path
  markdown = Get-Content -LiteralPath $resolvedMarkdownPath.Path -Raw -Encoding UTF8
  createdAt = (Get-Date).ToString("o")
}

$payload | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $pendingPath -Encoding utf8

Start-DevServerIfNeeded

Start-Process -FilePath $WordPath -ArgumentList "/w" | Out-Null
