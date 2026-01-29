# Copilot Instructions for pbi2xmla

This project is a PowerShell-based utility to extract a Power BI Desktop model and convert it into a TMSL (Tabular Model Scripting Language) script compatible with SQL Server Analysis Services (SSAS).

## Architecture & Data Flow
- **Single Component**: The entire logic resides in [PBI2XMLA.ps1](PBI2XMLA.ps1).
- **Data Flow**: Power BI Desktop (Local Analysis Services) -> Tabular Object Model (TOM) -> JSON Serialization -> Property Stripping/Transformation -> Output JSON.

## Critical Workflows
- **Running the Script**: Execute `./PBI2XMLA.ps1` from a PowerShell terminal.
- **Prerequisites**: 
  - Power BI Desktop must be running with a model loaded.
  - SQL Server Analysis Services client libraries (TOM) are loaded from the Power BI Desktop installation directory.

## Core Patterns & Conventions
- **General Purpose**: This script is intended to be a general-purpose utility. No model-specific workarounds or hardcoded logic for specific Power BI files are allowed.
- **TOM Interaction**: Uses `Microsoft.AnalysisServices.Tabular.Server` to connect to the local Power BI instance.
- **TMSL Manipulation**: 
    - The model is serialized to JSON, then converted to `[PSCustomObject]` for manipulation.
    - Power BI-specific properties (e.g., `lineageTag`, `annotations`, `variations`) are recursively removed using `Remove-Properties` to ensure compatibility with SSAS.
- **Partition Transformation**: 
    - Power BI uses 'm' (Power Query) partitions which are often rewritten by this script into 'query' (SQL) partitions for SSAS compatibility.
    - It attempts to extract SQL Server and Database names from M expressions using regex.

## Development Constraints
- **Library Paths**: Hardcoded to `C:\Program Files\Microsoft Power BI Desktop\bin\`. If libraries are missing, check if Power BI is installed in a different location.
- **Port Discovery**: Locates the dynamic port of Power BI via `$env:LOCALAPPDATA\Microsoft\Power BI *\AnalysisServicesWorkspaces\msmdsrv.port.txt`.
- **Output**: Writes the final JSON to `$env:USERPROFILE\Desktop\PowerBI_Model_Export.json`.

## Guidelines for AI
- **No Model-Specific Hacks**: Do not add logic that targets specific tables, columns, or measures from a particular model. All transformations must be generic and applicable to any Power BI model.
- When modifying JSON manipulation, maintain the recursive `Remove-Properties` pattern.
- Be cautious with Power BI model properties; many are invalid in SSAS and must be added to the `$propsToRemove` list if they cause deployment errors.
- Ensure any new dependency on .NET types is compatible with the version of Analysis Services libraries being loaded.
