# SolidWorks Release Pack

One-click add-in for SolidWorks that generates professional release packages — drawings, PDFs, DXFs, STEP files, BOMs, and preview images — from parts and assemblies.

## Features

- **Smart Drawing Generation** — auto-selects sheet size, places orthographic + isometric views, section & detail views for complex geometry
- **Assembly Drawings** — exploded views, auto-balloons, BOM tables
- **Multi-format Export** — PDF, DXF (flat-pattern for sheet metal), STEP, Parasolid
- **BOM Extraction** — Excel `.xlsx` with formatted output via ClosedXML
- **Version-Aware APIs** — runtime detection of SolidWorks version with automatic fallbacks (SW 2018+)
- **Feature Analysis** — 16-category feature classification for intelligent dimensioning
- **Batch Processing** — export current document, all children, or remote files

## Requirements

- SolidWorks 2018 or newer (API v26+)
- .NET Framework 4.8
- Windows x64

## Quick Start

1. **Build** the solution:
   ```
   dotnet build SolidWorksReleasePack.sln --configuration Release
   ```

2. **Register** the add-in (run as Administrator):
   ```
   scripts\register.bat
   ```

3. **Restart SolidWorks** → the add-in appears in *Tools → Add-Ins*

4. Open the **Release Pack** task pane, select outputs, and click **Generate**.

## Uninstall

```
scripts\unregister.bat
```

## Project Structure

```
src/
  ReleasePack.Engine/     Core logic (drawing gen, BOM, export, feature analysis)
  ReleasePack.AddIn/      SolidWorks COM add-in (TaskPane UI, toolbar)
scripts/
  register.bat            RegAsm registration
  unregister.bat          RegAsm unregistration
.github/workflows/
  build.yml               CI — build on push/PR
  release.yml             CD — package & release on version tag
```

## Releasing

Push a version tag to create a GitHub Release:
```
git tag v1.0.0
git push origin v1.0.0
```
