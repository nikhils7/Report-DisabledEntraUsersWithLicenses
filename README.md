
# Report Disabled Entra ID Users with Active Licenses

This PowerShell script identifies disabled users in Entra ID (Azure AD) who still have licenses assigned, generates a CSV report, and sends it via email using Microsoft Graph PowerShell SDK with certificate-based app-only authentication.

## 📋 Prerequisites

Before running the script, ensure the following are set up:

1. **Microsoft Graph PowerShell SDK** installed:
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```

2. **App Registration in Entra ID (Azure AD)**:
   - Register a new app in Azure Portal → Entra ID → App registrations.
   - Set account type to your organization only.

3. **Certificates**:
   - Upload a certificate to the app registration.
   - Install the same certificate on the machine running the script.
   - Note the certificate **Thumbprint**.

4. **API Permissions**:
   - Add the following Microsoft Graph permissions:
     - `User.Read.All`
     - `Directory.Read.All`
     - `Mail.Send`
   - Grant **Admin consent** for these permissions.

5. **Update Script Placeholders**:
   - `<YOUR-TENANT-ID>` → Found in Azure AD → Overview
   - `<YOUR-APP-CLIENT-ID>` → Found in App registration → Overview
   - `<YOUR-CERT-THUMBPRINT>` → From installed certificate
   - `<SENDER-EMAIL>` and `<RECIPIENT-EMAIL>` → Valid mailbox in tenant

## ⚙️ Configuration

Edit the following variables in the script:
```powershell
$TenantId   = '<YOUR-TENANT-ID>'
$ClientId   = '<YOUR-APP-CLIENT-ID>'
$CertThumb  = '<YOUR-CERT-THUMBPRINT>'

$Sender     = '<SENDER-EMAIL>'
$Recipient  = '<RECIPIENT-EMAIL>'
```

## 🚀 Usage

Run the script in PowerShell:
```powershell
.\DisabledUsersReport.ps1
```

## 📦 Output

- CSV file saved to `Documents\EntraReports`
- Email sent with CSV attached and HTML preview of first 10 rows

## 👤 Author

Nikhil Sawant – IT Senior Analyst

## 📝 License

This project is licensed under the MIT License. See the LICENSE file for details.
