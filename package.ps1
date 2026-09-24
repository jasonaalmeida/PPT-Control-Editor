# Creates a distributable ZIP package of the Shape Radius add-in

$projectRoot = $PSScriptRoot
$outputZip = Join-Path $projectRoot "ShapeRadius-Addin.zip"

# Remove old ZIP if it exists
if (Test-Path $outputZip) {
    Remove-Item $outputZip -Force
    Write-Host "Removed existing ZIP" -ForegroundColor Yellow
}

# Create temp staging directory
$stagingDir = Join-Path $env:TEMP "ShapeRadius-Package"
if (Test-Path $stagingDir) {
    Remove-Item $stagingDir -Recurse -Force
}
New-Item -ItemType Directory -Path $stagingDir | Out-Null

# Copy files to staging
Write-Host "Copying files..." -ForegroundColor Cyan
Copy-Item (Join-Path $projectRoot "manifest.xml") $stagingDir
Copy-Item (Join-Path $projectRoot "src") $stagingDir -Recurse
Copy-Item (Join-Path $projectRoot "README.md") $stagingDir
Copy-Item (Join-Path $projectRoot "assets") $stagingDir -Recurse -ErrorAction SilentlyContinue

# Update manifest to use relative paths (for local sideloading)
$manifestPath = Join-Path $stagingDir "manifest.xml"
$manifestContent = Get-Content $manifestPath -Raw
$manifestContent = $manifestContent -replace 'https://localhost:3000/', ''
Set-Content $manifestPath $manifestContent -NoNewline

Write-Host "Creating ZIP..." -ForegroundColor Cyan
Compress-Archive -Path "$stagingDir\*" -DestinationPath $outputZip -Force

# Cleanup
Remove-Item $stagingDir -Recurse -Force

Write-Host "`nPackage created: $outputZip" -ForegroundColor Green
Write-Host "Size: $([math]::Round((Get-Item $outputZip).Length / 1KB, 2)) KB" -ForegroundColor Gray
Write-Host "`nShare this ZIP file. Recipients should:" -ForegroundColor Yellow
Write-Host "  1. Extract to a permanent folder" -ForegroundColor White
Write-Host "  2. Share the folder (right-click > Properties > Sharing > Share)" -ForegroundColor White
Write-Host "  3. Note the UNC path (\\PC-NAME\FolderName)" -ForegroundColor White
Write-Host "  4. PowerPoint > File > Options > Trust Center > Trusted Add-in Catalogs" -ForegroundColor White
Write-Host "     -> paste UNC path -> Add catalog -> Show in Menu" -ForegroundColor White
Write-Host "  5. Restart PowerPoint" -ForegroundColor White
Write-Host "  6. Home -> Add-ins -> More Add-ins -> SHARED FOLDER tab -> 'Shape Radius'" -ForegroundColor White
Write-Host "`nSee README.md for full instructions and troubleshooting." -ForegroundColor Gray
