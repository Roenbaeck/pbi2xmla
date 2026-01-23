# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
# 1. Load the Analysis Services Libraries
try {
    Add-Type -Path "C:\Program Files\Microsoft Power BI Desktop\bin\Microsoft.AnalysisServices.Server.Tabular.dll"
    Add-Type -Path "C:\Program Files\Microsoft Power BI Desktop\bin\Microsoft.AnalysisServices.Server.Tabular.Json.dll" 
}
catch {
    Write-Error "Could not find TOM DLLs. Please install SQL Server client libraries or adjust the path."
    return
}

# 2. Find the Port of the running Power BI Desktop instance
# Power BI creates a file named 'msmdsrv.port.txt' in a temp directory for every running session.
$portFile = Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\Power BI *\AnalysisServicesWorkspaces" -Filter "msmdsrv.port.txt" -Recurse | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($null -eq $portFile) {
    Write-Error "Power BI Desktop does not appear to be running."
    return
}

$port = Get-Content $portFile.FullName -Encoding Unicode
$connectionString = "localhost:$port"

Write-Host "Connected to Power BI on port $port"

# 3. Connect to the internal Power BI Model
$server = New-Object Microsoft.AnalysisServices.Tabular.Server
$server.Connect($connectionString)

# There is usually only one database with a generic UUID name
$database = $server.Databases[0]

# 4. Serialize the Model to TMSL (JSON)
$options = New-Object Microsoft.AnalysisServices.Tabular.SerializeOptions
$options.IgnoreTimestamps = $true
$options.IgnoreInferredObjects = $true
$options.IgnoreInferredProperties = $true
$options.SplitMultilineStrings = $false

$script = [Microsoft.AnalysisServices.Tabular.JsonSerializer]::SerializeDatabase($database, $options)
$scriptObj = $script | ConvertFrom-Json

# Set name and compatibility level for SSAS 2022
$scriptObj.name = "PowerBIModel"
$scriptObj.compatibilityLevel = 1600

# Function to recursively remove specified properties from all objects
function Remove-Properties {
    param ($obj, [string[]]$properties)
    if ($obj -is [PSCustomObject]) {
        foreach ($prop in $properties) {
            $obj.PSObject.Properties.Remove($prop)
        }
        foreach ($objProp in @($obj.PSObject.Properties)) {
            if ($objProp.Value -is [PSCustomObject] -or $objProp.Value -is [Array]) {
                Remove-Properties -obj $objProp.Value -properties $properties
            }
        }
    } elseif ($obj -is [Array]) {
        foreach ($item in $obj) {
            Remove-Properties -obj $item -properties $properties
        }
    }
}

# Remove all problematic Power BI specific properties recursively
$propsToRemove = @(
    'lineageTag', 
    'changedProperties', 
    'annotations', 
    'modifiedTime', 
    'structureModifiedTime', 
    'refreshedTime', 
    'sourceProviderType', 
    'attributeHierarchy', 
    'summarizeBy', 
    'partitions', 
    'extendedProperties', 
    'expressions', 
    'queryGroups', 
    'state', 
    'relyOnReferentialIntegrity',
    'defaultPowerBIDataSourceVersion',
    'discourageImplicitMeasures',
    'dataAccessOptions',
    'sourceQueryCulture',
    'createdTimestamp',
    'lastUpdate',
    'lastSchemaUpdate',
    'lastProcessed',
    'dataSources',
    'cultures'
)
Remove-Properties -obj $scriptObj -properties $propsToRemove

# Remove rowNumber columns and dataType from measures
foreach ($table in $scriptObj.model.tables) {
    if ($table.columns) {
        $table.columns = @($table.columns | Where-Object { -not $_.PSObject.Properties['type'] -or $_.type -ne 'rowNumber' })
    }
    if ($table.measures) {
        foreach ($measure in $table.measures) {
            $measure.PSObject.Properties.Remove('dataType')
        }
    }
    if ($table.PSObject.Properties['type']) {
        $table.PSObject.Properties.Remove('type')
    }
    
    # Add partition for each table (required by SSAS)
    $table | Add-Member -NotePropertyName 'partitions' -NotePropertyValue @(
        @{
            name = $table.name
            mode = 'import'
            source = @{
                type = 'query'
                query = "SELECT * FROM [$($table.name)]"
                dataSource = 'SqlServer localhost'
            }
        }
    ) -Force
}

# Add a placeholder data source
$scriptObj.model | Add-Member -NotePropertyName 'dataSources' -NotePropertyValue @(
    @{
        name = 'SqlServer localhost'
        connectionString = 'Provider=SQLNCLI11;Data Source=localhost;Initial Catalog=YourDatabase;Integrated Security=SSPI'
        impersonationMode = 'impersonateServiceAccount'
    }
) -Force

# Create the TMSL command
$scriptWrapped = @{
    createOrReplace = @{
        object = @{
            database = "PowerBIModel"
        }
        database = $scriptObj
    }
}

# Convert to JSON
$jsonOutput = $scriptWrapped | ConvertTo-Json -Depth 100

# 5. Output to file
$outputPath = "$env:USERPROFILE\Desktop\PowerBI_Model_Export.json"
$jsonOutput | Out-File $outputPath -Encoding UTF8

Write-Host "Success! Model definition exported to $outputPath"
Write-Host "You can now open this JSON in SSMS and execute it against your SSAS Tabular server."
Write-Host ""
Write-Host "NOTE: Update the data source connection string before running on SSAS."
