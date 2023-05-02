Connect-PnPOnline -Url "https://{TenantName}.sharepoint.com/sites/{site}
#Update Viva connection in teams.
Publish-PnPCompanyApp -PortalUrl 'https://{TenantName}.sharepoint.com' -AppName 'Portal' -ColoredIconPath 'C:\temp\onlineSuper-Charger-192px.png' -OutlineIconPath 'C:\temp\DrinkingFountain32px.png' -AccentColor '#FF0000' -Force

#Connect to Teams
Connect-MicrosoftTeams

#Get existing Pinned apps.
$GlobalSettings = Get-CsTeamsAppSetupPolicy -Identity Global
$PinnedApps = $GlobalSettings.PinnedAppBarApps
$PinnedAppBarApps = @()
foreach($PinnedApp in $PinnedApps)
{
  $app = New-Object -TypeName Microsoft.Teams.Policy.Administration.Cmdlets.Core.PinnedApp -Property @{Id="$($PinnedApp.Id)"}
  $PinnedAppBarApps.Add($app)
}

#Pinned new app.
Set-CsTeamsAppSetupPolicy -Identity 'Set-Test' -PinnedAppBarApps $PinnedAppBarApps
