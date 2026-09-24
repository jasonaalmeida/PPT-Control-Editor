# Generates manifest-prod.xml from manifest.xml with URLs pointing at your
# hosted files (GitHub Pages, Netlify, etc.) and a fresh add-in GUID so
# it doesn't collide with the dev version in PowerPoint's cache.
#
# Usage:
#   .\build-prod-manifest.ps1 -BaseUrl "https://YOUR-USERNAME.github.io/YOUR-REPO"
#
# Example (for GitHub Pages on repo "PPT-Control-Editor" owned by "jason-almeida"):
#   .\build-prod-manifest.ps1 -BaseUrl "https://jason-almeida.github.io/PPT-Control-Editor"
#
# The generated manifest-prod.xml is what you sideload into PowerPoint.

param(
    [Parameter(Mandatory=$true)]
    [string]$BaseUrl
)

# Strip trailing slash if present
$BaseUrl = $BaseUrl.TrimEnd('/')

# Validate: must be HTTPS (Office Add-ins require it in production)
if (-not $BaseUrl.StartsWith('https://')) {
    Write-Error "BaseUrl must start with https:// (Office Add-ins require HTTPS in production)."
    exit 1
}

$projectRoot = $PSScriptRoot
$sourceManifest = Join-Path $projectRoot "manifest.xml"
$outputManifest = Join-Path $projectRoot "manifest-prod.xml"

if (-not (Test-Path $sourceManifest)) {
    Write-Error "manifest.xml not found at $sourceManifest"
    exit 1
}

Write-Host "Reading manifest.xml..." -ForegroundColor Cyan
$content = Get-Content $sourceManifest -Raw

# Replace every localhost:3000 URL with the hosted base URL
$content = $content -replace 'https://localhost:3000', $BaseUrl

# Generate a fresh GUID so PowerPoint treats this as a separate add-in
# from any dev version you have installed
$newGuid = [guid]::NewGuid().ToString()
Write-Host "Assigning new add-in GUID: $newGuid" -ForegroundColor Cyan

# Replace the <Id> element's value (first occurrence only)
$content = $content -replace '(?<=<Id>)[0-9a-fA-F-]+(?=</Id>)', $newGuid

# Also update the AppDomain (just the origin, no path)
$origin = ([System.Uri]$BaseUrl).GetLeftPart([System.UriPartial]::Authority)
$content = $content -replace '<AppDomain>https://[^<]+</AppDomain>', "<AppDomain>$origin</AppDomain>"

Set-Content -Path $outputManifest -Value $content -NoNewline -Encoding UTF8

Write-Host "`nCreated: $outputManifest" -ForegroundColor Green
Write-Host "  Base URL: $BaseUrl" -ForegroundColor Gray
Write-Host "  Add-in GUID: $newGuid" -ForegroundColor Gray
Write-Host "`nNext steps:" -ForegroundColor Yellow
Write-Host "  1. Make sure your files are pushed to GitHub and Pages is enabled" -ForegroundColor White
Write-Host "  2. Verify the URL loads in a browser: $BaseUrl/src/taskpane/taskpane.html" -ForegroundColor White
Write-Host "  3. Sideload manifest-prod.xml (either via a network share, or via" -ForegroundColor White
Write-Host "     M365 admin center centralized deployment)" -ForegroundColor White
