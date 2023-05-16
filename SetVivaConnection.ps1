Param(    [parameter(Mandatory = $false)] [Switch]$ConfigureVivaConnection = $false,
    [parameter(Mandatory = $false)] [Switch]$ProvisionTeamsApp = $true
)
$LogsDirectoryName = "C:\Diksha\scripts"
$LogFileName = "$LogsDirectoryName\AutomatedDeployment_" + $(get-date -f yyyyMMdd) + ".log"
#region Logging Functions

Function CheckLogsDirectory {

    if (-not (Test-Path $LogsDirectoryName)) {

        New-Item -ErrorAction Ignore -ItemType directory -Path $LogsDirectoryName | Out-Null
    }
}

Function DeleteOldLogs {

    try {

        if ($PurgeLogDays -gt 0) {

            Write-Log "Deleting Log files older than $PurgeLogDays days" -Level Info
            $currentDate = Get-Date
            $dateToDelete = $CurrentDate.AddDays(-$PurgeLogDays)
            Get-ChildItem $LogsDirectoryName | Where-Object { $_.CreationTime -lt $DatetoDelete } | Remove-Item
            Write-Log "Log files deleted" -Level Info
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
        [ValidateSet("Error", "Warn", "Info", "Finish", "Start")]
        [string]$Level = "Info",

        [Parameter(Mandatory = $false)]
        [switch]$NoClobber
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
                Write-Host $Message	-ForegroundColor Yellow
                $LevelText = 'WARNING'
            }
            'Info' {
                Write-Host $Message
                $LevelText = 'INFO'
            }
            'Start' {
                Write-Host $Message -ForegroundColor Cyan
                $LevelText = 'START'
            }
            'Finish' {
                Write-Host $Message -ForegroundColor Green
                $LevelText = 'FINISH'
            }
        }

        # Write log entry to $Path
        "$FormattedDate $LevelText`t$Message" | Out-File -FilePath $LogFileName -Append
    }
    End {
    }
}

#endregion Logging Functions

#region Private Functions
function ConfigureVivaConnections($configuration)
{
    try
    {
        Write-Log "Viva connection configuration started..." -Level Info
        #Configuring viva connection.
        Publish-PnPCompanyApp -PortalUrl $configuration.VivaConnections.PortalUrl -AppName $configuration.VivaConnections.AppName -CompanyName $configuration.VivaConnections.CompanyName -CompanyWebSiteUrl $configuration.VivaConnections.CompanyWebSiteUrl -ColoredIconPath "C:\Temp\portal.png" -OutlineIconPath "C:\Temp\portal1.png" -AccentColor '#FF0000' -Description 'View the Intranet portal in a Teams App (Viva Connections)' -LongDescription 'View the Intranet portal in a Teams App using Viva Connections and we can keep on writing a longer and longer description here.'  -Force


    }
    catch
    {
        Write-Log "Error while configuring viva connection. Error details: $($_.Exception.Message)" -Level Error
    }
}

function ProvisionTeamsApp($configuration)
{
    try
    {
        if($configuration.Teams.Apps.App.Filename -ne $null)
        {
            #$customApp = New-TeamsApp -DistributionMethod "organization" -Path $configuration.Teams.Apps.App.Filename
        }
        
        foreach($currPolicy in $configuration.Teams.SetupPolicies.SetupPolicy)
        {
            TeamsPolicySettings -CustomAppId "0aa7553d-dc89-47e3-9ec2-30867afc335d"<#$customApp.Id#> -configuration $configuration.Teams.SetupPolicies.SetupPolicy
        }
    }
    catch
    {
        Write-Log "Error while provsioning teams app. Error details: $($_.Exception.Message)" -Level Error
    }
}
function TeamsPolicySettings($CustomAppId, $configuration)
{
    try
    {
         $pinnedApps = @()
        $customApp = New-Object -TypeName Microsoft.Teams.Policy.Administration.Cmdlets.Core.PinnedApp   -Property @{Id="$($CustomAppId)"}

        #Getting Global settings.
        $GlobalSettings = Get-CsTeamsAppSetupPolicy -Identity $configuration.Name

        $pinnedApps = $GlobalSettings.PinnedAppBarApps
        #$pinnedApps = $GlobalSettings.PinnedMessageBarApps
        $pinnedApps+=$customApp
        #$pinnedApps.Add($customApp)
        #$GlobalSettings.PinnedAppBarApps+=$customApp
        #$pinnedApps = $GlobalSettings.PinnedAppBarApps
        #$GlobalSettings.PinnedAppBarApps.Add($customApp)
        $pinnedApps.Length
        $pinnedApps
        #Pin new app.
        #Set-CsTeamsAppSetupPolicy -Identity "Test Policy" -PinnedAppBarApps $pinnedApps
        Set-CsTeamsAppSetupPolicy -Identity "Global" -PinnedAppBarApps $pinnedApps

        #Assign Permission to User(s).
        foreach($currUser in $configuration.Users.User)
        {
            Write-Log "Assigning permission on $($configuration.Name) to $($currUser.LoginName)" -Level Info
            Grant-CsTeamsAppPermissionPolicy -Identity $currUser.LoginName -PolicyName $configuration.Name
        }

        #Assign Permission to Group(s).
        if($configuration.Name -ne "Global")
        {
            foreach($currGroup in $configuration.Groups.Group)
            {
                Write-Log "Assigning permission on $($configuration.Name) to $($currGroup.LoginName)" -Level Info
                New-CsGroupPolicyAssignment -GroupId $currGroup.LoginName -PolicyName $configuration.Name -Rank 1 –PolicyType "TeamsAppSetupPolicy"
            }            
        }
    }
    catch
    {
        Write-Log "Error while setting policy for $($configuration.Name). Error details: $($_.Exception.Message)" -Level Error
    }
}
#endregion


#region Main Process
function MainProcess()
{
    [xml]$config = Get-Content -Path "C:\Diksha\scripts\Config.xml"

        Write-Log "Connecting to $($config.Configuration.VivaConnections.PortalUrl) site."
        Connect-PnPOnline -Url $config.Configuration.VivaConnections.PortalUrl -Interactive
        Write-Log "Connecting to Microsoft teams using azure login..."
        Connect-MicrosoftTeams
    

    if($ConfigureVivaConnection){
        Write-Log "Configuring viva connection..."
        ConfigureVivaConnections -configuration $config.Configuration
        Write-Log "Viva connection configured..."
    }

    if($ProvisionTeamsApp){
        Write-Log "Provisioning Teams App..."
        ProvisionTeamsApp -configuration $config.Configuration
        Write-Log "Teams app Provisioned..."
    }
}
#endregion
#Install-Module -Name MicrosoftTeams -Force -AllowClobber
CheckLogsDirectory
DeleteOldLogs
MainProcess
