<#
.SYNOPSIS
  Deploys Rhino Forms to a SharePoint Online tenant and sets up a test site.

.DESCRIPTION
  1. Adds rhino-forms.sppkg to the tenant App Catalog and deploys it tenant-wide.
  2. Installs the app on the target site.
  3. Uploads the sample config to SiteAssets/RhinoForms/ProjectTracker.json.
  4. Creates two full-page app pages: RhinoDesigner.aspx and RhinoList.aspx.

  Requires PnP.PowerShell 2.x (Install-Module PnP.PowerShell -Scope CurrentUser)
  and an account that is a SharePoint Administrator on the tenant.

.EXAMPLE
  ./deploy/Deploy-RhinoForms.ps1 -SiteUrl https://1ykyqr.sharepoint.com

.EXAMPLE
  ./deploy/Deploy-RhinoForms.ps1 -SiteUrl https://1ykyqr.sharepoint.com/sites/RhinoTest -ClientId <your-entra-app-id>
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory)] [string] $SiteUrl,
  [string] $PackagePath = (Join-Path $PSScriptRoot '..\sharepoint\solution\rhino-forms.sppkg'),
  [string] $SampleConfigPath = (Join-Path $PSScriptRoot '..\samples\project-tracker.json'),
  # PnP.PowerShell 2.12+ requires your own Entra app registration for interactive login.
  # Create one with: Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "PnP PowerShell" -Tenant <tenant>.onmicrosoft.com
  [string] $ClientId,
  [switch] $SkipPages
)

$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable PnP.PowerShell)) {
  throw 'PnP.PowerShell is not installed. Run: Install-Module PnP.PowerShell -Scope CurrentUser'
}
if (-not (Test-Path $PackagePath)) {
  throw "Package not found at $PackagePath. Build it first: npx gulp bundle --ship; npx gulp package-solution --ship"
}
$PackagePath = (Resolve-Path $PackagePath).Path
$SampleConfigPath = (Resolve-Path $SampleConfigPath).Path

$tenantHost = ([Uri]$SiteUrl).Host
$adminUrl = "https://$($tenantHost -replace '\.sharepoint\.com$', '-admin.sharepoint.com')"

function Connect-Site([string] $Url) {
  if ($ClientId) { Connect-PnPOnline -Url $Url -Interactive -ClientId $ClientId }
  else { Connect-PnPOnline -Url $Url -Interactive }
}

# ---------------------------------------------------------------------------
Write-Host "==> Connecting to tenant admin centre $adminUrl" -ForegroundColor Cyan
Connect-Site $adminUrl

$catalogUrl = Get-PnPTenantAppCatalogUrl
if (-not $catalogUrl) {
  $catalogUrl = "https://$tenantHost/sites/appcatalog"
  Write-Host "==> No tenant App Catalog found. Creating $catalogUrl (this can take a few minutes)" -ForegroundColor Yellow
  $owner = (Get-PnPConnection).PSCredential.UserName
  if (-not $owner) { $owner = Read-Host 'Owner UPN for the App Catalog site' }
  Register-PnPAppCatalogSite -Url $catalogUrl -Owner $owner -TimeZoneId 76   # 76 = Canberra/Melbourne/Sydney
}
Write-Host "==> App Catalog: $catalogUrl" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
Write-Host "==> Uploading and deploying $([IO.Path]::GetFileName($PackagePath))" -ForegroundColor Cyan
Connect-Site $catalogUrl
$app = Add-PnPApp -Path $PackagePath -Scope Tenant -Overwrite -Publish -SkipFeatureDeployment:$false
Write-Host "    App id: $($app.Id)  version: $($app.AppCatalogVersion)"

# ---------------------------------------------------------------------------
Write-Host "==> Installing on $SiteUrl" -ForegroundColor Cyan
Connect-Site $SiteUrl
$installed = Get-PnPApp -Identity $app.Id -Scope Tenant
if (-not $installed.InstalledVersion) {
  Install-PnPApp -Identity $app.Id -Scope Tenant -Wait
}
elseif ($installed.InstalledVersion -ne $installed.AppCatalogVersion) {
  Write-Host "    Updating from $($installed.InstalledVersion) to $($installed.AppCatalogVersion)"
  Update-PnPApp -Identity $app.Id -Scope Tenant
}
else {
  Write-Host "    Already installed (v$($installed.InstalledVersion))"
}

# ---------------------------------------------------------------------------
Write-Host "==> Uploading sample config to SiteAssets/RhinoForms" -ForegroundColor Cyan
$assets = Get-PnPList -Identity 'Site Assets' -ErrorAction SilentlyContinue
if (-not $assets) {
  # Provisions the Site Assets library if the site does not have one yet.
  Invoke-PnPSPRestMethod -Method Post -Url '/_api/web/lists/EnsureSiteAssetsLibrary' | Out-Null
  $assets = Get-PnPList -Identity 'Site Assets'
}
Resolve-PnPFolder -SiteRelativePath 'SiteAssets/RhinoForms' | Out-Null
Add-PnPFile -Path $SampleConfigPath -Folder 'SiteAssets/RhinoForms' -NewFileName 'ProjectTracker.json' | Out-Null
$configFileUrl = "$($assets.RootFolder.ServerRelativeUrl)/RhinoForms/ProjectTracker.json"
Write-Host "    $configFileUrl"

# ---------------------------------------------------------------------------
if (-not $SkipPages) {
  Write-Host "==> Creating pages" -ForegroundColor Cyan

  # Designer page (full-page app)
  Add-PnPPage -Name 'RhinoDesigner' -Title 'Rhino List Designer' -LayoutType SingleWebPartAppPage -Publish -ErrorAction SilentlyContinue | Out-Null
  Add-PnPPageWebPart -Page 'RhinoDesigner' -Component 'Rhino List Designer' | Out-Null
  Set-PnPPage -Identity 'RhinoDesigner' -Publish | Out-Null
  Write-Host "    $SiteUrl/SitePages/RhinoDesigner.aspx"

  # List page (full-page app) bound to the sample config
  Add-PnPPage -Name 'RhinoList' -Title 'Project Tracker' -LayoutType SingleWebPartAppPage -Publish -ErrorAction SilentlyContinue | Out-Null
  Add-PnPPageWebPart -Page 'RhinoList' -Component 'Rhino List' -WebPartProperties @{
    configSource  = 'file'
    configFileUrl = $configFileUrl
    hideHeader    = $false
  } | Out-Null
  Set-PnPPage -Identity 'RhinoList' -Publish | Out-Null
  Write-Host "    $SiteUrl/SitePages/RhinoList.aspx"
}

Write-Host ''
Write-Host 'Done.' -ForegroundColor Green
Write-Host "Next: open $SiteUrl/SitePages/RhinoDesigner.aspx, choose Open > From Site Assets > ProjectTracker.json,"
Write-Host "then click 'Provision list'. The RhinoList page will light up once the list exists."
