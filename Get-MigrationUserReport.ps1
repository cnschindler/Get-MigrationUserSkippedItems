[cmdletbinding()]
##
Param(
[Parameter(Mandatory=$true)]
[Parameter(ParameterSetName="Detect")]
[System.IO.DirectoryInfo]$BasePath
)

$LogFolderName = "Get-MigrationUserReport"
[string]$LogPath = Join-Path -Path $BasePath -ChildPath $LogFolderName
[string]$LogfileFullPath = Join-Path -Path $LogPath -ChildPath ("Get-MigrationUserReport_{0:yyyyMMdd-HHmmss}.log" -f [DateTime]::Now)
$Script:NoLogging
[string]$CSVFullPath = Join-Path -Path $LogPath -ChildPath ("Get-MigrationUserSkippedItems_{0:yyyyMMdd-HHmmss}.txt" -f [DateTime]::Now)

function Write-LogFile
{
    # Logging function, used for progress and error logging...
    # Uses the globally (script scoped) configured LogfileFullPath variable to identify the logfile and NoLogging to disable it.
    #
    [CmdLetBinding()]

    param
    (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [System.Management.Automation.ErrorRecord]$ErrorInfo = $null
    )
    # Prefix the string to write with the current Date and Time, add error message if present...

    if ($ErrorInfo)
    {
        $logLine = "{0:d.M.y H:mm:ss} : [Error] : {1}: {2}" -f [DateTime]::Now, $Message, $ErrorInfo.Exception.Message
    }

    else
    {
        $logLine = "{0:d.M.y H:mm:ss} : [INFO] : {1}" -f [DateTime]::Now, $Message
    }

    if (!$Script:NoLogging)
    {
        # Create the Script:Logfile and folder structure if it doesn't exist
        if (-not (Test-Path $Script:LogfileFullPath -PathType Leaf))
        {
            New-Item -ItemType File -Path $Script:LogfileFullPath -Force -Confirm:$false -WhatIf:$false | Out-Null
            Add-Content -Value "Logging started." -Path $Script:LogfileFullPath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        }

        # Write to the Script:Logfile
        Add-Content -Value $logLine -Path $Script:LogfileFullPath -Encoding UTF8 -WhatIf:$false -Confirm:$false
        Write-Verbose $logLine
    }
    else
    {
        Write-Host $logLine
    }
}
function Confirm-EXOModuleInstalled
{
    $EXOModuleInstalled = Get-Module -ListAvailable -Name "ExchangeOnlineManagement" | Sort-Object Version -Descending | Select-Object -First 1

    if ($EXOModuleInstalled.Name -eq "ExchangeOnlineManagement")
    {
        $EXOConnection = Get-ConnectionInformation

        if ($EXOConnection.State -ne "Connected")
        {
            Connect-ExchangeOnline
        }
    }

    else
    {
        Write-Host -ForegroundColor Red -Object "Exchange Online Management Module not installed. Please install the module and connect to exchange online before using this script."
        Exit
    }

}

Confirm-EXOModuleInstalled

# Build a Datatable
$Datatable = New-Object System.Data.DataTable
$Properties = @("Kind","ScoringClassifications","FolderName","Subject")

foreach ($property in $Properties)
{
    $Datatable.Columns.Add($property) | Out-Null
}

$Datatable.Columns.Add("Identity")

$MigrationUsersWithSkippedItems = Get-MigrationUser -Status Synced | Where-Object {$_.SkippedItemCount -gt 0}

ForEach ($User in $MigrationUsersWithSkippedItems)
{
    $SkippedItems = (Get-MigrationUserStatistics -Identity $User.Identity -IncludeSkippedItems).SkippedItems | Select-Object $Properties

    foreach ($Item in $SkippedItems)
    {
        $row = $Datatable.NewRow()
        $row.Identity = $User.Identity
        $row.Kind = $item.Kind
        $row.ScoringClassifications = $item.ScoringClassifications
        $row.FolderName = $item.FolderName
        $row.Subject = $item.Subject
        $Datatable.Rows.Add($row)
    }
}

$Datatable | Out-GridView
