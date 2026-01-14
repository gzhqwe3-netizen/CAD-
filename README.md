# RebarCadSync

AutoCAD .NET plugin that syncs rebar annotations in DWG from an Excel source of truth.

> **Safety**: Always save a DWG backup before running `REBARSYNC`.

## Requirements

- AutoCAD with .NET API assemblies.
- .NET Framework 4.8 SDK or Visual Studio Build Tools.
- Environment variable `ACAD_DIR` pointing to the AutoCAD install folder that contains:
  - `acmgd.dll`
  - `acdbmgd.dll`

Example (PowerShell):

```powershell
$env:ACAD_DIR = "C:\Program Files\Autodesk\AutoCAD 2024"
```

## Project Structure

```
src/
  RebarCadSync/
    RebarCadSync.csproj
    RebarSync.cs
    Properties/
      AssemblyInfo.cs
build.ps1
README.md
```

## Build

```powershell
./build.ps1
```

The build fails with a clear error if `ACAD_DIR` is missing or does not contain the required DLLs.

## Load in AutoCAD

1. Build the project.
2. In AutoCAD, run `NETLOAD`.
3. Browse to `src\RebarCadSync\bin\Release\RebarCadSync.dll` (or `Debug` if built that way).
4. Run the command: `REBARSYNC`.

## Excel Template

File: `.xlsx`

Sheet name: `Rebar`

Columns:

| mark | qty | dia | length_cm |
|------|-----|-----|-----------|

- `mark`: circled mark text (e.g. `1`, `3'`, `3a`)
- `qty`: quantity (integer)
- `dia`: diameter (integer)
- `length_cm`: length in centimeters (number)

## DWG Update Behavior

- Scans **ModelSpace** `TEXT` and `MTEXT`.
- Mark text must be a `DBText` on layer **Dim**.
- For each mark found in DWG that exists in Excel:
  - Find nearest **spec** text within radius **350** drawing units that matches regex `^\s*\d+\s*HA\s*\d+\s*$`.
  - Find nearest **length** text within radius **350** drawing units that matches regex `^\s*\d+(\.\d+)?\s*$` and **does not** match the spec regex.
  - Update spec to `${qty}HA${dia}` (no spaces).
  - Update length to `${length_cm}` formatted as integer if near integer else one decimal.

## Logging

A CSV log is written to the Desktop after running `REBARSYNC`:

- Summary counts
- Per-mark changes
- Marks not found or ambiguous matches

