# SolidWorks Release Pack — Version Compatibility

## Current Build Configuration

| Setting                        | Value                  |
|-------------------------------|------------------------|
| **Interop NuGet Version**      | 32.1.0 (SW 2024 API)  |
| **Target Framework**           | .NET Framework 4.8     |
| **Platform**                   | x64                    |
| **Min Recommended SW Version** | 2018 (API v26)         |

## SolidWorks Version → API Version Map

| SW Version | API Version | Interop NuGet Version | Status              |
|-----------|-------------|----------------------|---------------------|
| SW 2018   | 26.x        | 26.x.0               | ⚠ Needs testing    |
| SW 2019   | 27.x        | 27.x.0               | ⚠ Needs testing    |
| **SW 2020** | **28.x**  | **28.x.0**            | **🎯 Target**       |
| SW 2021   | 29.x        | 29.x.0               | ⚠ Needs testing    |
| SW 2022   | 30.x        | 30.x.0               | ⚠ Needs testing    |
| SW 2023   | 31.x        | 31.x.0               | ⚠ Needs testing    |
| SW 2024   | 32.x        | 32.1.0                | ✅ Compiles (current)|

## Known Compatibility Risks (v32 → v28 downgrade)

The following API calls in our code may not exist in SW 2020. Each is wrapped in
try/catch so the add-in won't crash, but the feature may silently fail.

### High Risk — May need fallback

| File | Method / API Call | Risk | Fallback |
|------|-------------------|------|----------|
| `DrawingGenerator.cs` | `CreateSectionViewAt5` | Added post-2020; may not exist | Catches exception, logs warning |
| `DrawingGenerator.cs` | `CreateDetailViewAt4` | Signature may differ in v28 | Catches exception, logs warning |
| `DrawingGenerator.cs` | `Create3rdAngleViews2` | Should exist in v28 | Falls back to `PlaceViewsManually` |
| `AssemblyDrawingGenerator.cs` | `AutoBalloon4` | May have fewer params in v28 | Catches exception, logs warning |
| `AssemblyDrawingGenerator.cs` | `InsertBomTable4` | Param count may differ | Catches exception, logs warning |
| `DimensionEngine.cs` | `InsertModelAnnotations3` | Should exist in v28+ | ✅ Safe |
| `ExportPipeline.cs` | `Extension.SaveAs3` | Should exist in v28+ | ✅ Safe |
| `ExportPipeline.cs` | `ExportToDWG2` | Should exist in v28+ | ✅ Safe |
| `DependencyScanner.cs` | `OpenDoc6` | Exists since v24+ | ✅ Safe |

### Low Risk — Core API stable since v24+

| API | Status |
|-----|--------|
| `Feature.GetTypeName2()` | ✅ Stable |
| `Feature.GetDefinition()` | ✅ Stable |
| `IWizardHoleFeatureData2` | ✅ Stable |
| `ISimpleFilletFeatureData2` | ✅ Stable |
| `IExtrudeFeatureData2` | ✅ Stable |
| `CustomPropertyManager.Add3` | ✅ Stable |
| `ModelDoc2.Extension.SelectByID2` | ✅ Stable |
| `PartDoc.GetPartBox` | ✅ Stable |
| `AssemblyDoc.GetComponents` | ✅ Stable |

## How to Switch to SW 2020 Interop

To compile against the exact SW 2020 API (recommended for your setup):

### Step 1: Update NuGet packages in both `.csproj` files

```xml
<!-- Change from -->
<PackageReference Include="SolidWorks.Interop.sldworks" Version="32.1.0" />
<PackageReference Include="SolidWorks.Interop.swconst" Version="32.1.0" />

<!-- Change to (SW 2020) -->
<PackageReference Include="SolidWorks.Interop.sldworks" Version="28.2.0" />
<PackageReference Include="SolidWorks.Interop.swconst" Version="28.2.0" />
```

Files to edit:
- `src/ReleasePack.Engine/ReleasePack.Engine.csproj`
- `src/ReleasePack.AddIn/ReleasePack.AddIn.csproj`

### Step 2: Restore & rebuild

```powershell
dotnet restore
dotnet build SolidWorksReleasePack.sln
```

### Step 3: Fix any new errors

If methods/enums are missing in v28, they were added after SW 2020.
The code already has try/catch guards, so most will degrade gracefully.
Document any compile errors here for reference.

## Verification Checklist for SW 2020

After switching to v28 interop, test these features:

- [ ] Add-in loads in SW 2020 (appears in Tools → Add-Ins)
- [ ] CommandManager button appears in toolbar
- [ ] TaskPane opens correctly
- [ ] Current Document scope works
- [ ] Current + Children scope works
- [ ] Drawing generation — standard views placed
- [ ] Drawing generation — section view (if applicable)
- [ ] Drawing generation — detail view (if applicable)
- [ ] PDF export works
- [ ] DXF export works (sheet metal flat pattern)
- [ ] STEP export works
- [ ] BOM Excel export works
- [ ] Assembly drawing — auto-balloons
- [ ] Assembly drawing — BOM table in drawing
- [ ] Title block populated correctly
- [ ] Smart dimensioning — model annotations inserted
- [ ] Smart dimensioning — fillet/chamfer notes
- [ ] Smart dimensioning — thread callouts
- [ ] Smart dimensioning — pattern notes

## Build Error Log

_Record any compile errors when switching interop versions here:_

### v28.2.0 (SW 2020) — Not yet tested
```
(pending)
```

### v32.1.0 (SW 2024) — ✅ Builds successfully
```
Build succeeded with 5 warnings (all CS0649 unassigned struct fields in DimensionEngine.cs)
Date: 2026-02-19
```
