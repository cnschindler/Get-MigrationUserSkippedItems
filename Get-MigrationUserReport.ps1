[cmdletbinding()]
##
Param(
[Parameter(Mandatory=$true)]
[Parameter(ParameterSetName="Detect")]
[System.IO.DirectoryInfo]$BasePath,
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

$MigrationUsersWithSkippedItems = Get-MigrationUser -Status Synced | Where-Object {$_.SkippedItemCount -gt 0}

