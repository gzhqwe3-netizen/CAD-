Param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"

if (-not $env:ACAD_DIR) {
    Write-Error "ACAD_DIR environment variable is not set. Please set ACAD_DIR to your AutoCAD install folder (containing acmgd.dll and acdbmgd.dll)."
}

$projectPath = Join-Path $PSScriptRoot "src\RebarCadSync\RebarCadSync.csproj"

Write-Host "Building RebarCadSync ($Configuration)..."

& msbuild $projectPath /t:Restore,Build /p:Configuration=$Configuration
