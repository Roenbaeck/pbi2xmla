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
# This creates the "Create Database" script
$options = [Microsoft.AnalysisServices.Tabular.SerializeOptions]::Default
$script = [Microsoft.AnalysisServices.Tabular.JsonSerializer]::SerializeDatabase($database, $options)
$scriptObj = $script | ConvertFrom-Json
$scriptObj.name = "PowerBIModel"
$scriptWrapped = @{
    create = @{
        database = $scriptObj
    }
} | ConvertTo-Json -Depth 100

# 5. Output to file
$outputPath = "$env:USERPROFILE\Desktop\PowerBI_Model_Export.json"
$scriptWrapped | Out-File $outputPath

Write-Host "Success! Model definition exported to $outputPath"
Write-Host "You can now open this JSON in SSMS and execute it against your SSAS Tabular server."
