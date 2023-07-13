<#

.SYNOPSIS
	This script provides you all the active sites with Administrator and Storage used..

.DESCRIPTION


.Example    
    .\SiteInventory.ps1 

.NOTES
Pre-requisites:
    Install-Module -Name PnP.PowerShell
    
    

#>

#region GlobalVairable
$LogsDirectoryName = "Logs"
$LogFileName = ".\$LogsDirectoryName\SiteInventory_" + $(get-date -f MMddyyyyHHmmss) + ".log"
$PurgeLogDays = 30
#endregion
#region Logging Functions
Function CheckLogDirectory {
    if (-not (Test-Path $LogsDirectoryName)) {
        New-Item -ErrorAction Ignore -ItemType directory -Path $LogsDirectoryName
    }
} 

Function DeleteOldLogs {

    try {

        if ($PurgeLogDays -gt 0) {

            Write-Log "Deleting Log files older than $PurgeLogDays days"
            $currentDate = Get-Date
            $dateToDelete = $currentDate.AddDays(-$PurgeLogDays)
            Get-ChildItem $LogsDirectoryName | Where-Object { $_.CreationTime -lt $dateToDelete } | Remove-Item
            Write-Log "Log files deleted"
        }
    }
    catch {

        Write-Log "Error deleting Log files. Details: $($_.Exception.Message)" -Level Error
    }
}

Function Write-Log {

    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true)]
        [ValidateNotNullOrEmpty()]
        [Alias("LogContent")]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Error", "Warn", "Info", "Success")]
        [string]$Level = "Info"
    )

    Begin {
        # Set VerbosePreference to Continue so that verbose messages are displayed.
        $VerbosePreference = 'SilentlyContinue'
    }
    Process {
        # Format Date for our Log File
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Write message to error, warning, or verbose pipeline and specify $LevelText
        switch ($Level) {
            'Error' {
                Write-Host "ERROR: $Message" -ForegroundColor Red
                $LevelText = 'ERROR'
            }
            'Warn' {
                Write-Host $Message -ForegroundColor Yellow
                $LevelText = 'WARNING'
            }
            'Info' {
                Write-Host $Message
                $LevelText = 'INFO'
            }
            'Success' {
                Write-Host $Message -f Green
                $LevelText = 'Success'
            }
        }
        # Write log entry to $Path
        "$FormattedDate $LevelText`t$Message" | Out-File -FilePath $LogFileName -Append
    }
    End {
    }
}
#endregion

#region Private Functions
function MainProcess()
{
    try
    {
        Write-Log "Connecting to SharePoint Admin" -Level Info
        # Connect to SharePoint Online
        Connect-PnPOnline -Url "https://pc45-admin.sharepoint.com/" -Interactive
        Write-Log "Connected" -Level Success

        # Get all SharePoint sites
        $sites = Get-PnPTenantSite
        
        # Create an array to store the results
        $results = @()
        
        # Iterate through each site and gather required information
        foreach ($site in $sites) {
            $siteUrl = $site.Url
            Write-Log "Processing for site with url: $($siteUrl)" -Level Info
            Connect-PnPOnline -Url $siteUrl -Interactive
            # Get site administrators
            $admins = Get-PnPSiteCollectionAdmin | Select-Object -ExpandProperty Title
        
            # Get site storage size
            $storageSize = Get-PnPTenantSite -Url $siteUrl | Select-Object -ExpandProperty StorageUsageCurrent
        
            # Create a custom object with the site information
            $siteInfo = [PSCustomObject]@{
                SiteUrl = $siteUrl
                Administrators = $admins -join ";"
                StorageSize = $storageSize.ToString() +" MB(s)"
            }
        
            # Add the site information to the results array
            $results += $siteInfo
        }
        
        # Output the results as a CSV file
        $results | Export-Csv -Path "SiteInventory.csv" -NoTypeInformation
        
        # Disconnect from SharePoint Online
        Disconnect-PnPOnline
    }
    catch
    {
    Write-Log "Error while fetching data: Error Details: $($_.Exception)" -Level Error
    }
}
#endregion







CheckLogDirectory
DeleteOldLogs
MainProcess
