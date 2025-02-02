<#
    .SYNOPSIS
        Extracts patient data from ORBBI Dent database and exports to CSV.
    
    .DESCRIPTION
        Reads connection details from Dental.log file, connects to the ORBBI Dent Access database,
        extracts patient information and exports it to a CSV file with UTF-8 encoding.
    
    .PARAMETER OrbbiDataPath
        The path to the folder containing ORBBI Dent data files (Dental.log and database file).
        Default value is "C:\OrbbiData"
    
    .PARAMETER OutputCsvPath
        The path where the CSV file with patient data will be saved.
        Default value is "C:\OrbbiData\patients.csv"
    
    .OUTPUTS
        CSV file containing patient records with the following columns:
        - First_Name
        - Family_Name
        - Id
        - Phone
        - Email
        - District
        - Address
    
    .EXAMPLE
        .\Read-OrbbiDentDatabase.ps1
        Exports patient data using default paths
    
    .EXAMPLE
        .\Read-OrbbiDentDatabase.ps1 -OrbbiDataPath "D:\ORBBI" -OutputCsvPath "D:\export.csv"
        Exports patient data from custom location to specified output file
    
    .NOTES
        Requirements:
        - Microsoft Access Database Engine (32 or 64 bit depending on PowerShell version)
        - Read access to ORBBI Dent database
        - Write permission to output location
#>
[CmdletBinding()]
param (
    [Parameter(HelpMessage = "The path to the OrbbiData folder")]
    [String] $OrbbiDataPath = "C:\OrbbiData",

    [Parameter(HelpMessage = "Path tothe output CSV file")]
    [String] $OutputCsvPath = "C:\OrbbiData\patients.csv"
)

$Encoding = [System.Text.Encoding]::UTF8
[System.Console]::InputEncoding = $Encoding
[Console]::OutputEncoding = $Encoding

$LogFilePath = "$OrbbiDataPath\Dental.log"

$DbConnectionStringRegex = '\[.+\] .+Data Source="(?<datasource>.+)".+Password=(?<pwd>.+);.+'
$RegexMatches = (Get-Content -Path $LogFilePath -ErrorAction 'Stop' `
    | Select-String -Pattern $DbConnectionStringRegex)[0].Matches[0]
$DbDatasource = $RegexMatches.Groups['datasource'].Value
if (-not $DbDatasource) {
    $DbDatasource = Join-Path -Path $OrbbiDataPath -ChildPath 'p.dent'
    Write-Warning "Could not find datasource in '$LogFilePath'. Defaulting to '$DbDatasource'."
}
$DbDatasource = Join-Path -Path $OrbbiDataPath -ChildPath (Split-Path -Path $DbDatasource -Leaf)

$DbPwd = $RegexMatches.Groups['pwd'].Value
if (-not $DbPwd) {
    $DbPwd = 'B&chochi&B'
    Write-Warning "Could not find database password in '$LogFilePath'. Defaulting to '$DbPwd'."
}
"Database connection pwd:`n$DbPwd"

$DbProvider = if ([Environment]::Is64BitProcess) { "Microsoft.ACE.OLEDB.12.0" } else { "Microsoft.Jet.OLEDB.4.0" }
$DbConnectionString = "Provider=$DbProvider;Data Source=""$DbDatasource"";Jet OLEDB:Database Password=$DbPwd;"

"Database connection string:`n$DbConnectionString"

if (-not $DbConnectionString) {
    Write-Error "Could not find the database connection string in the Dental.log file."
    return
}

Add-Type -AssemblyName System.Data

$Connection = New-Object System.Data.OleDb.OleDbConnection($DbConnectionString)
try {
    $Connection.Open() | Out-Null
    $Command = $Connection.CreateCommand()
    $Command.CommandText = "SELECT * FROM patients WHERE 1;"
    $Reader = $Command.ExecuteReader()

    $Result = @()
    while ($Reader.Read()) {
        $Result += [ordered]@{
            First_Name = $Reader['Given']
            Family_Name = $Reader['Family']
            Id = $Reader['identifier']
            Phone = $Reader['Phone']
            Email = $Reader['Email']
            District = $Reader['kvartal']
        }
        $AddressParts = @();
        if ('' -ne $Reader['ulica'] -and '' -ne $Reader['ulnomer']) {
            $AddressParts += "ул. $($Reader['ulica']) $($Reader['ulnomer'])"
        }
        if ('' -ne $Reader['blok']) {
            $AddressParts += "бл. $($Reader['blok'])"
        }
        if ('' -ne $Reader['vhod']) {
            $AddressParts += "вх. $($Reader['vhod'])"
        }
        if ('' -ne $Reader['ap']) {
            $AddressParts += "ап. $($Reader['ap'])"
        }
        $Result[-1].Address = $AddressParts -join ', '
    }
    "Found $($Result.Count) patients."
}
finally {
    if ($Reader) {
        $Reader.Close()
    }
    if ($Connection) {
        $Connection.Close()
    }
}

$Result | ForEach-Object {[PSCustomObject] $_} `
    | ConvertTo-Csv -NoTypeInformation `
    | Out-File -FilePath $OutputCsvPath -Encoding 'UTF8'

Write-Host "Exported patient data to '$OutputCsvPath'." -ForegroundColor Green
