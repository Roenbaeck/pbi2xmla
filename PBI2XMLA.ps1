# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
param(
    [string]$SqlServerOverride = ""
)
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
Write-Host "Model contains $($database.Model.DataSources.Count) data sources."

# 4. Serialize the Model to TMSL (JSON)
$options = New-Object Microsoft.AnalysisServices.Tabular.SerializeOptions
$options.IgnoreTimestamps = $true
$options.IgnoreInferredObjects = $false
$options.IgnoreInferredProperties = $false
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
$propsToRemove = @('lineageTag', 'changedProperties', 'annotations', 'modifiedTime', 'structureModifiedTime', 'refreshedTime', 'sourceProviderType', 'attributeHierarchy', 'summarizeBy', 'extendedProperties', 'queryGroups', 'state', 'relyOnReferentialIntegrity', 'defaultPowerBIDataSourceVersion', 'discourageImplicitMeasures', 'dataAccessOptions', 'sourceQueryCulture', 'createdTimestamp', 'lastUpdate', 'lastSchemaUpdate', 'lastProcessed', 'cultures', 'variations')
Remove-Properties -obj $scriptObj -properties $propsToRemove

# --- CRITICAL: Link M Partitions to Data Sources ---
$modelDataSources = @()
if ($scriptObj.model.PSObject.Properties['dataSources']) { 
    $modelDataSources = @($scriptObj.model.dataSources) 
}

if ($modelDataSources.Count -eq 0) {
    # If no data sources found (common in PBI Desktop), we MUST create one for SSAS to authorize M queries
    $dsName = "SqlServer localhost"
    $connStr = "Provider=MSOLEDBSQL;Data Source=localhost;Initial Catalog=YourDatabase;Integrated Security=SSPI"
    
    # Try to extract actual connection info from the first partition's M code
    foreach ($table in $scriptObj.model.tables) {
        if ($table.partitions -and $table.partitions[0].source.expression) {
            $m = $table.partitions[0].source.expression
            if ($m -match 'Sql\.Database\("([^"]+)",\s*"([^"]+)"\)') {
                $srv = $matches[1]
                $db = $matches[2]
                if ([string]::IsNullOrWhiteSpace($SqlServerOverride)) {
                    $dsName = "SqlServer $srv $db"
                    $connStr = "Provider=MSOLEDBSQL;Data Source=$srv;Initial Catalog=$db;Integrated Security=SSPI"
                } else {
                    $dsName = "SqlServer $SqlServerOverride $db"
                    $connStr = "Provider=MSOLEDBSQL;Data Source=$SqlServerOverride;Initial Catalog=$db;Integrated Security=SSPI"
                }
                break
            }
        }
    }
    
    $newDS = @{
        name = $dsName
        connectionString = $connStr
        impersonationMode = "impersonateServiceAccount"
    } | ConvertTo-Json | ConvertFrom-Json
    
    if (-not $scriptObj.model.PSObject.Properties['dataSources']) {
        $scriptObj.model | Add-Member -NotePropertyName "dataSources" -NotePropertyValue @($newDS) -Force
    } else {
        $scriptObj.model.dataSources = @($newDS)
    }
    $modelDataSources = @($newDS)
    Write-Host "Created missing Data Source: $dsName"
}

if ($modelDataSources.Count -gt 0) {
    $primaryDSName = $modelDataSources[0].name
    foreach ($table in $scriptObj.model.tables) {
        if ($table.PSObject.Properties['partitions']) {
            foreach ($partition in $table.partitions) {
                if ($partition.PSObject.Properties['source']) {
                    # SSAS does not accept M partitions. Convert M to query partitions.
                    if ($partition.source.PSObject.Properties['type'] -and $partition.source.type -eq 'm') {
                        $schema = 'dbo'
                        $item = $table.name
                        if ($partition.source.PSObject.Properties['expression']) {
                            $m = $partition.source.expression
                            if ($m -match '\[Schema\s*=\s*"([^"]+)",\s*Item\s*=\s*"([^"]+)"\]') {
                                $schema = $matches[1]
                                $item = $matches[2]
                            }
                        }
                        $partition.source = @{
                            type = 'query'
                            query = "SELECT * FROM [$schema].[$item]"
                            dataSource = $primaryDSName
                        } | ConvertTo-Json | ConvertFrom-Json
                        if ($partition.PSObject.Properties['dataSource']) {
                            $partition.PSObject.Properties.Remove('dataSource')
                        }
                    }
                    # For query partitions, dataSource belongs inside source
                    elseif ($partition.source.PSObject.Properties['type'] -and $partition.source.type -eq 'query') {
                        $partition.source | Add-Member -MemberType NoteProperty -Name "dataSource" -Value $primaryDSName -Force
                    }
                }
            }
        }
    }
}

# Clean up Data Sources for SSAS
if ($modelDataSources.Count -gt 0) {
    foreach ($ds in $modelDataSources) {
        $ds.PSObject.Properties.Remove('credential')
        $ds.PSObject.Properties.Remove('privacyLevel')
        if ($ds.PSObject.Properties['connectionDetails']) { $ds.PSObject.Properties.Remove('connectionString') }
        if (-not [string]::IsNullOrWhiteSpace($SqlServerOverride) -and $ds.PSObject.Properties['connectionString']) {
            $ds.connectionString = $ds.connectionString -replace 'Data Source=([^;]+)', "Data Source=$SqlServerOverride"
        }
    }
}

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
}

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
