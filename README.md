# pbi2xmla

`pbi2xmla` is a PowerShell utility designed to extract a Power BI Desktop model and convert it into a Tabular Model Scripting Language (TMSL) script compatible with SQL Server Analysis Services (SSAS) Tabular.

## Features

- **Automated Port Discovery**: Automatically finds the local Analysis Services port used by a running Power BI Desktop instance.
- **PBI-to-SSAS Transformation**: Cleanses the model by removing Power BI-specific properties (e.g., `lineageTag`, `annotations`, `variations`) that are invalid in SSAS.
- **Partition Conversion**: Attempts to convert 'm' (Power Query) partitions into 'query' (SQL) partitions for broader SSAS compatibility.
- **Data Source Generation**: Automatically creates a SQL Server data source definition if one is missing.

## Prerequisites

- **Power BI Desktop**: Must be installed in the default location (`C:\Program Files\Microsoft Power BI Desktop\`).
- **Running Model**: Power BI Desktop must be open with the model you wish to export loaded.
- **PowerShell**: Execution policy must allow script execution (e.g., `RemoteSigned`).

## Usage

1. Open a PowerShell terminal.
2. Navigate to the project directory.
3. Run the script:
   ```powershell
   .\PBI2XMLA.ps1
   ```
4. **Optional**: Use the `-SqlServerOverride` parameter to specify a different SQL Server for the data sources:
   ```powershell
   .\PBI2XMLA.ps1 -SqlServerOverride "my-sql-server"
   ```

## Output

The script generates a file named `PowerBI_Model_Export.json` on your **Desktop**. This JSON file contains a `createOrReplace` TMSL command. 

To deploy the model to SSAS:
1. Open **SQL Server Management Studio (SSMS)**.
2. Connect to your **Analysis Services** Tabular instance.
3. Open a new **XMLA Query** window (**New Query** -> **XMLA**).
4. Paste the content of the generated JSON file and execute it.

## Crucial Note on Mixed Mode Models

While Power BI supports **Mixed Mode** models (combining DirectQuery and Import storage modes), **SSAS Tabular does not support this configuration**. 

Models using a mix of DirectQuery and Import partitions will likely fail to deploy or function correctly in SSAS. This script is intended for single-mode models, and manual adjustments to the exported JSON may be required for complex scenarios.

## Important Considerations

- **Library Dependencies**: The script loads Tabular Object Model (TOM) libraries directly from the Power BI Desktop installation directory.
- **DirectQuery**: Basic DirectQuery models are supported, but the script primarily focuses on converting M-based import partitions to SQL queries.
- **Cleanup**: After deployment to SSAS, you may still need to update data source credentials and connection strings in SSMS.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
