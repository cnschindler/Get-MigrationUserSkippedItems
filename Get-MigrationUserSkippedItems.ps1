[cmdletbinding()]

Param(
[Parameter(Mandatory=$true)]
[System.IO.DirectoryInfo]$BasePath
)

$LogFolderName = "Get-MigrationUserSkippedItemsReport"
[string]$LogPath = Join-Path -Path $BasePath -ChildPath $LogFolderName
[string]$LogfileFullPath = Join-Path -Path $LogPath -ChildPath ($LogFolderName + "_{0:yyyyMMdd-HHmmss}.log" -f [DateTime]::Now)
$Script:NoLogging
[string]$CSVFullPath = Join-Path -Path $LogPath -ChildPath ("MigrationUserSkippedItems_{0:yyyyMMdd-HHmmss}.txt" -f [DateTime]::Now)

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
$Datatable.Columns.Add("Identity")
$Properties = @("Kind","ScoringClassifications","FolderName","Subject")
foreach ($property in $Properties)
{
    $Datatable.Columns.Add($property) | Out-Null
}

# Get all Migrationusers with skipped items
$MigrationUsersWithSkippedItems = Get-MigrationUser -Status Synced | Where-Object {$_.SkippedItemCount -gt 0}

$Message = "Found $($MigrationUsersWithSkippedItems.Count) migrationusers with skipped items."
Write-Host -ForegroundColor Green -Object $Message
Write-LogFile -Message $Message

# Loop through result
ForEach ($User in $MigrationUsersWithSkippedItems)
{
    $Message = "Currently processing user $($user.Identity)"
    Write-Host -ForegroundColor Green -Object $Message
    Write-LogFile -Message $Message

    # Get skipped items of current migrationuser    
    $SkippedItems = (Get-MigrationUserStatistics -Identity $User.Identity -IncludeSkippedItems).SkippedItems | Select-Object $Properties

    $Message = "Found $($SkippedItems.Count) skipped items"
    Write-Host -ForegroundColor Green -Object $Message
    Write-LogFile -Message $Message

    # Add each skipped item to the datatable
    $Message = "Adding item details to datatable"
    Write-Host -ForegroundColor Green -Object $Message
    Write-LogFile -Message $Message

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

#Export the datable to CSV
$Message = "Exporting datatable to CSV file $($CSVFullPath)"
Write-Host -ForegroundColor Magenta -Object $Message
Write-LogFile -Message $Message

try
{
    $Datatable | Export-Csv -Path $CSVFullPath -NoTypeInformation -ErrorAction Stop
    $Message =  "CSV file successfully written."
    Write-Host -ForegroundColor Green -Object $Message
    Write-LogFile -Message $Message
}

catch
{
    $Message = "Error writing CSV file."
    Write-Host -ForegroundColor Red -Object ($Message + " $($_)")
    Write-LogFile -Message $Message -ErrorInfo $_
}
