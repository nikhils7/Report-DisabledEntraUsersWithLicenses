<#
.SYNOPSIS
  Reports all disabled Entra ID users who still have licenses assigned and emails a CSV from the same mailbox.

.NOTES
  - Uses Microsoft Graph PowerShell SDK with certificate-based app-only auth.
  - Skips re-import if Microsoft.Graph.Authentication is already loaded to avoid assembly conflicts.
  - ASCII-only to prevent encoding/smart character issues.

  ⚠️ IMPORTANT:
  Replace the following placeholders before running:
    - <YOUR-TENANT-ID>
    - <YOUR-APP-CLIENT-ID>
    - <YOUR-CERT-THUMBPRINT>
    - <SENDER-EMAIL>
    - <RECIPIENT-EMAIL>
#>

# ---------------------- Config ----------------------
$TenantId   = '<YOUR-TENANT-ID>'
$ClientId   = '<YOUR-APP-CLIENT-ID>'
$CertThumb  = '<YOUR-CERT-THUMBPRINT>'

$Sender     = '<SENDER-EMAIL>'      # from mailbox (same as recipient)
$Recipient  = '<RECIPIENT-EMAIL>'   # to mailbox

$ReportFolder = Join-Path $env:USERPROFILE 'Documents\EntraReports'
$ReportName   = 'DisabledUsersWithLicenses_{0:yyyy-MM-dd_HHmm}.csv' -f (Get-Date)
$ReportPath   = Join-Path $ReportFolder $ReportName

$LogPath      = Join-Path $ReportFolder ('RunLog_{0:yyyy-MM-dd}.log' -f (Get-Date))
# ----------------------------------------------------

# Create folder and start transcript
New-Item -Path $ReportFolder -ItemType Directory -Force | Out-Null
Start-Transcript -Path $LogPath -Append -ErrorAction SilentlyContinue | Out-Null

# Make installs non-interactive
$ProgressPreference = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

function Ensure-GraphModules {
    Write-Host "Checking Microsoft Graph module..."

    # If Authentication submodule is already loaded, skip re-import to avoid assembly conflicts.
    $mgAuthLoaded = Get-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
    if ($mgAuthLoaded) {
        Write-Host "Microsoft.Graph.Authentication already loaded (v$($mgAuthLoaded.Version)). Skipping re-import."
        return
    }

    # Ensure PSGallery and NuGet are available (silently)
    try {
        if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
            Register-PSRepository -Default
        }
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    } catch {
        Write-Warning "Repository setup warning: $($_.Exception.Message)"
    }

    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
    } catch {
        Write-Warning "NuGet provider warning: $($_.Exception.Message)"
    }

    # Install minimal submodules if missing
    $modulesToInstall = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Mail'
    )

    foreach ($m in $modulesToInstall) {
        if (-not (Get-Module -ListAvailable -Name $m -ErrorAction SilentlyContinue)) {
            try {
                Write-Host "Installing $m ..."
                Install-Module $m -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -ErrorAction Stop
            } catch {
                throw "Failed to install ${m}: $($_.Exception.Message)"
            }
        }
    }

    # Import submodules (no RequiredVersion to avoid cross-version conflicts)
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Users, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Mail -ErrorAction SilentlyContinue

    Write-Host "Microsoft Graph modules ready."
}

function Get-CertificateByThumbprint {
    param(
        [Parameter(Mandatory)][string]$Thumbprint
    )
    $paths = @("Cert:\CurrentUser\My\$Thumbprint", "Cert:\LocalMachine\My\$Thumbprint")
    foreach ($p in $paths) {
        $c = Get-Item -Path $p -ErrorAction SilentlyContinue
        if ($c) { return $c }
    }
    throw "Certificate with thumbprint $Thumbprint not found in CurrentUser\My or LocalMachine\My."
}

try {
    Ensure-GraphModules

    # Connect to Graph with certificate (app-only)
    $cert = Get-CertificateByThumbprint -Thumbprint $CertThumb
    Write-Host "Connecting to Microsoft Graph (app-only cert auth)..."
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -Certificate $cert -NoWelcome -ErrorAction Stop

    # Select profile if available; otherwise default (v1.0) is fine
    if (Get-Command Select-MgProfile -ErrorAction SilentlyContinue) {
        Select-MgProfile -Name 'v1.0'
    }

    $ctx = $null
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        $ctx = Get-MgContext
    }
    if ($ctx) {
        Write-Host "Connected to tenant: $($ctx.TenantId) as app: $($ctx.ClientId)"
    } else {
        Write-Host "Connected to Microsoft Graph."
    }

    # Get SKU map for license name resolution
    Write-Host "Fetching subscribed SKUs..."
    $skus = Get-MgSubscribedSku -All -ErrorAction Stop
    $skuMap = @{}
    foreach ($s in $skus) {
        $skuMap[$s.SkuId.ToString()] = $s.SkuPartNumber
    }

    # Query disabled users and requested properties
    Write-Host "Querying disabled users..."
    $count = 0
    $users = Get-MgUser `
        -All `
        -ConsistencyLevel eventual `
        -CountVariable count `
        -Filter "accountEnabled eq false" `
        -Property "id,displayName,userPrincipalName,accountEnabled,assignedLicenses" `
        -ErrorAction Stop

    Write-Host "Disabled users returned: $count"

    # Filter to those with assigned licenses
    $disabledWithLicenses = $users | Where-Object {
        $_.AssignedLicenses -and $_.AssignedLicenses.Count -gt 0
    }
    Write-Host "Disabled users WITH licenses: $($disabledWithLicenses.Count)"

    # Prepare CSV rows
    $rows = foreach ($u in $disabledWithLicenses) {
        $licenseNames = foreach ($lic in $u.AssignedLicenses) {
            $key = $lic.SkuId.ToString()
            if ($skuMap.ContainsKey($key)) { $skuMap[$key] } else { $key }
        }
        [pscustomobject]@{
            DisplayName        = $u.DisplayName
            UserPrincipalName  = $u.UserPrincipalName
            ObjectId           = $u.Id
            AccountEnabled     = $u.AccountEnabled
            LicenseSkuPartNos  = ($licenseNames | Sort-Object -Unique) -join '; '
            LicenseCount       = ($licenseNames | Measure-Object).Count
        }
    }

    # Export CSV
    Write-Host "Exporting CSV to $ReportPath ..."
    $rows | Sort-Object UserPrincipalName | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8

    # Compose email
    $now = Get-Date
    $totalDisabled = $count
    $withLicenses  = $disabledWithLicenses.Count
    $previewList   = ($rows | Select-Object -First 10 | ForEach-Object {
        "<tr><td>$($_.DisplayName)</td><td>$($_.UserPrincipalName)</td><td>$($_.LicenseCount)</td><td>$($_.LicenseSkuPartNos)</td></tr>"
    }) -join "`n"

    $subject = "Disabled Entra ID users with active licenses - $($now.ToString('yyyy-MM-dd HH:mm zzz'))"

$bodyHtml = @"
<p>Hi Team,</p>
<p>Here is the report of <b>disabled Entra ID users who still have licenses assigned</b>.</p>
<ul>
  <li>Total disabled users (in tenant): <b>$totalDisabled</b></li>
  <li>Disabled users with licenses: <b>$withLicenses</b></li>
</ul>
<p>The full CSV is attached. Preview of first 10 rows:</p>
<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;">
  <thead>
    <tr><th>Name</th><th>UPN</th><th>License Count</th><th>Licenses</th></tr>
  </thead>
  <tbody>
    $previewList
  </tbody>
</table>
<p>Generated: $now</p>
"@

    $attachment = @{
        '@odata.type' = '#microsoft.graph.fileAttachment'
        Name          = [IO.Path]::GetFileName($ReportPath)
        ContentType   = 'text/csv'
        ContentBytes  = [Convert]::ToBase64String([IO.File]::ReadAllBytes($ReportPath))
    }

    $message = @{
        Subject      = $subject
        Body         = @{
            ContentType = 'HTML'
            Content     = $bodyHtml
        }
        ToRecipients = @(@{ EmailAddress = @{ Address = $Recipient } })
        Attachments  = @($attachment)
    }

    Write-Host "Sending email from $Sender to $Recipient with attachment..."
    Send-MgUserMail -UserId $Sender -Message $message -SaveToSentItems -ErrorAction Stop

    Write-Host "Report sent successfully."
}
catch {
    Write-Error $_.Exception.Message
    Write-Error $_.Exception | Format-List -Force
}
finally {
    try {
        if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
            $gctx = Get-MgContext
            if ($gctx) { Disconnect-MgGraph | Out-Null }
        }
    } catch {
        # swallow disconnect errors
    }
    try {
        Stop-Transcript | Out-Null
    } catch {
        # swallow transcript stop errors
    }
}
