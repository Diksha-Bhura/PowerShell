<#

.SYNOPSIS
	This script sets default values to it based on csv data.

.DESCRIPTION


.Example    
    .\SetDefaultTaggingValue.ps1 -CSVFilePath "[CSV File Path]" 

.NOTES
Pre-requisites:
    We need To install Nightly Version (2.0.28-nightly), To install we need to run below commands 
    Uninstall-Module -Name PnP.PowerShell  -AllVersions
    Install-Module -Name PnP.PowerShell -RequiredVersion 2.0.28-nightly -AllowPrerelease -AllowClobber
    
    
.LINK
    https://github.com/pnp/pnpframework/pull/840
    https://github.com/pnp/pnpframework/issues/838
#>

<#Param(
    [parameter(mandatory = $true)][string]$CSVFilePath
)#>

Import-Module PnP.PowerShell

#region GlobalVairable
$LogsDirectoryName = "Logs"
$LogFileName = ".\$LogsDirectoryName\MGB.SetDocumentTaggingFieldsValue_" + $(get-date -f MMddyyyyHHmmss) + ".log"
$PurgeLogDays = 30
$DocumentLibraryName = "Documents"
$defaultNull = "null"
$global:ErrorMessage = ""
$CSVFilePath="C:\Diksha\BSTFS\MGB\ProvisionDocumentLibraryTagging\DefaultValue.csv"
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
function GetTermString($termID) {
    $termValue = ""

    try {
        if ($termID -ne "" -and $null -ne $termID -and $termID -ne $defaultNull) {
            $term = Get-PnPTerm -Identity $termID
            $termName = $term.Name
            $termValue = "1;#$termName|$termID"
        }
        else {
            Write-Log "Term ID is blank" -Level Info
        }
    }
    catch {
        Write-Log "Error getting Term by ID: $termID. Details: $($_.Exception.Message)" -Level Error
    }

    return $termValue
}

function GetPnPConnection($siteURL) {
    $pnpConnection = Connect-PnPOnline -Url $siteURL -Interactive 
    return $pnpConnection
}

function SetTaggingValue($taggingData)
{
    $siteUrl = $taggingData.SiteUrl
    try
    {
        Write-Log "Connecting to site $($SiteCollectionUrl)" -Level Info
        $pnpConnection = GetPnPConnection -siteURL $SiteCollectionUrl 
        Write-Log "Connected to site $($SiteCollectionUrl)"  -Level Success

        #Getting tagging values.
        $departmentTermID = GetTermString $taggingData.bsoneDepartment
        $informationTopicsTermId = GetTermString $taggingData.bsoneInformationTopics

        $itemToUpdate = Get-PnPListItem -List $SitePagesLibraryName -Id $taggingData.ItemId

        Write-Log "Working on updating $($taggingData.ItemId)"

        $ctx = Get-PnPContext
        $ctx.Load($itemToUpdate.File)
        $ctx.ExecuteQuery()

        if($itemToUpdate.File.CheckOutType -ne "Online"){
            Set-PnPListItem -List $SitePagesLibraryName -Identity $taggingData.ItemId -Values 
@{"bsoneLegalEntity"=$legalEntityTermID; "bsoneInformationChannels" = $informationChannelTermID; "bsoneInformationType" = $informationTypeTermID; "bsoneDepartment" = $deptartmentTermID; "bsoneInformationTopics" = $informationTopicsTermID; "bsoneBreadcrumbParent" = $breadcrumbParentTermID;} -UpdateType SystemUpdate
        }
        else{
            Write-Log "File is checked out. So cannot update it."
        }

        Disconnect-PnPOnline
    }
    catch
    {
    }
}
#endregion


function MainProcess() {

    Write-Log "Importing CSV from $($CSVFilePath)"
    $taggingData = Import-Csv -Path $CSVFilePath

    if($taggingData -ne $null)
    {
        Write-Log "CSV data found."

        foreach($currenttaggingData in $taggingData)
        {
            SetTaggingValue -taggingData $currenttaggingData
            
        }

    }
}

CheckLogDirectory
DeleteOldLogs
MainProcess
$global:ErrorMessage
