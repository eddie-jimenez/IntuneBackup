# IntuneBackup Solution

Automated Intune backup solution to be run inside an Azure Automation Runbook.

## 📋 Overview

This runbook performs the following:

- Authenticates to Microsoft Graph using **Managed Identity**
- Backs up key Microsoft Intune configuration categories (no module dependency)
- Organizes backup files by category and app platform (Windows, macOS, etc.)
- Uploads extracted files to a **SharePoint Online** folder  
  `https://<your-tenant>.sharepoint.com/sites/<YourSite>/Shared Documents/Intune Backups/YYYY/MM`
- Sends detailed HTML summary reports (Success or Failure) to specified recipients
- Fully scheduled via **Azure Automation** to run monthly

---

## 🔧 How It Works

- Uses native `Invoke-RestMethod` with Graph endpoints to pull data
- Backs up each category into a timestamped temp folder
- Mobile apps are sorted by platform using the app’s `@odata.type`
- Uploads all extracted files (not zipped) to SharePoint via Microsoft Graph
- Sends HTML-formatted email status reports from a configurable sender address
- Cleans up local temp files after execution

---

## 🗂️ Backed-Up Categories

| Category                       | Graph Endpoint |
|--------------------------------|----------------|
| Device Configurations          | `/deviceConfigurations` |
| Compliance Policies            | `/deviceCompliancePolicies` |
| Configuration Policies         | `/configurationPolicies?$expand=settings` |
| Device Scripts                 | `/deviceManagementScripts` |
| App Policies & Configs         | `/mobileApps`, `/managedAppPolicies`, `/mobileAppConfigurations`, `/policySets` |
| Autopilot Profiles             | `/windowsAutopilotDeploymentProfiles`, `/deviceEnrollmentConfigurations` |
| Settings Catalog               | `/configurationPolicies?$expand=settings` |
| Feature/Quality/Driver Updates | `/windowsFeatureUpdateProfiles`, etc. |
| Conditional Access Policies    | `/identity/conditionalAccess/policies`, etc. |
| Mac Scripts & Custom Attributes| `/deviceShellScripts`, `/deviceCustomAttributeShellScripts` |
| Miscellaneous                  | Notification Templates, Role Tags, Terms & Conditions, Intune Branding, etc. |

> ❌ **Groups** are intentionally excluded to avoid backing up thousands of unnecessary AAD objects.

---

### ➕ Modifying Backup Categories

To **add or remove** a category from the backup logic:

1. Locate the `$categories` hashtable inside the script:

    ```powershell
    $categories = @{
        "DeviceConfigurations" = "/deviceConfigurations"
        "CompliancePolicies"   = "/deviceCompliancePolicies"
        ...
    }
    ```

2. To **add a category**, insert a new key-value pair using the correct Graph endpoint:

    ```powershell
    "EnrollmentStatusPageProfiles" = "/deviceEnrollmentConfigurations"
    ```

3. To **remove a category**, delete or comment out its entry:

    ```powershell
    # "ManagedAppPolicies" = "/managedAppPolicies"
    ```

4. If the category requires query parameters such as `$expand` or `$filter`, include them in the value:

    ```powershell
    "MobileApps" = "/deviceAppManagement/mobileApps?\$expand=assignments,categories"
    ```

> 💡 All updates will apply during the next scheduled or manual backup run.  
> ⚠️ Make sure each endpoint is valid and that the Azure Automation Managed Identity has the necessary Graph permissions.

---

## 🕒 Schedule

The runbook can be configured to run whenever works best for you.

For example:

> 🗓 **1st of every month at 10:00 AM EST**

This schedule can be adjusted in the Azure Automation account under **Schedules**.

---

## 📧 Email Reports

Emails include:

- ✅ Backup Success status  
- ❌ Backup Failure with error details  
- ⏱️ Runtime duration  
- 📁 SharePoint link to uploaded backup  
- 📊 Category item counts  
- ⚠️ Failure reports with troubleshooting steps

  <img width="518" height="628" alt="image" src="https://github.com/user-attachments/assets/a97963b9-c1bb-46ae-9833-e00e950a5794" />
  <img width="469" height="1213" alt="image" src="https://github.com/user-attachments/assets/6d28c0b3-4925-4f33-9460-ac8591b85bf6" />



---

## 🔧 Usage

Before using this runbook in your environment, review and replace the following placeholders in the script:

| Placeholder                    | Description                                                                                                                            |
|--------------------------------|----------------------------------------------------------------------------------------------------------------------------------------|
| `YOUR SITE URI`                | Full Microsoft Graph API URI for your SharePoint site. Example: `https://graph.microsoft.com/v1.0/sites/{tenant}.sharepoint.com:/sites/{site-name}` |
| `YOUR SHAREPOINT PATH`         | Folder path inside the SharePoint document library where backups are uploaded. Example: `ITadmins/Intune Backups`                 |
| `YOUR SHAREPOINT URL`          | Base SharePoint site web URL used to construct clickable folder links in email reports. Example: `https://{tenant}.sharepoint.com/sites/{site-name}/` |
| `YOUR TEST EMAIL`              | Default fallback email recipient for test runs. Update this to your own address or distribution list. Example: `'you@example.com'`    |
| `YOUR AUTOMATION ACCOUNT EMAIL`| Email address used to send status emails from Azure Automation. Must be a valid Entra ID user. Example: `automation@yourdomain.com` |


> ✅ If you're running this inside **Azure Automation**, only `YOUR SITE URI`, `YOUR SHAREPOINT PATH`, `YOUR SHAREPOINT URL`, and `'YOUR TEST EMAIL'` need to be updated.  
> ⚠️ Local testing requires valid Microsoft Graph authentication using `Connect-MgGraph` with a certificate or client secret.

## 🧪 Testing

You can test this script **locally** by providing custom parameters — but note that **authentication is required** since Managed Identity is only available in Azure-hosted environments.

### 🔐 Required for Manual Testing

You must authenticate using an **App Registration** in Entra ID (Azure AD) with one of the following:
> Note: This logic is not baked into the script. 

#### ✅ Recommended: Certificate Authentication

```powershell
Connect-MgGraph -ClientId "<your-client-id>" `
                -TenantId "<your-tenant-id>" `
                -CertificateThumbprint "<your-cert-thumbprint>"
```

#### ⚠️ Temporary (Not Secure): Client Secret Authentication

```powershell
Connect-MgGraph -ClientId "<your-client-id>" `
                -TenantId "<your-tenant-id>" `
                -ClientSecret "<your-client-secret>"
```

ℹ️ Make sure the App Registration has all necessary Microsoft Graph API permissions.
> See Required Permissions below...

---

### 🧪 Test Parameters

Once authenticated, you can run the backup manually with:

```powershell
param (
    [string]$EmailRecipient = "you@example.com"
)
```

Override the `$EmailRecipient` value if you'd like to receive a test HTML summary email.

---

## 🔐 Required Permissions

The Azure Automation Managed Identity or App Registration must have the following Microsoft Graph API permissions:

```
DeviceManagementConfiguration.Read.All  
DeviceManagementApps.Read.All  
DeviceManagementRBAC.Read.All  
DeviceManagementManagedDevices.Read.All  
Mail.Send  
Sites.ReadWrite.All  
Directory.Read.All  
Policy.Read.All
```

---

## 📚 Resources

- [Microsoft Graph API Documentation](https://learn.microsoft.com/graph/api/overview)
- [Managed Identity Auth for Azure Automation](https://learn.microsoft.com/azure/automation/enable-managed-identity)
- [Microsoft Graph REST API Permissions](https://learn.microsoft.com/graph/permissions-reference)
- [PowerShell - Upload to SharePoint via Graph](https://learn.microsoft.com/sharepoint/dev/apis/upload-large-files)

---

## 🛠️ Troubleshooting

See the failure email report for built-in diagnostics, or use these common checks:

- ✅ **Authentication** – Verify Managed Identity or App Registration is authorized
- 📦 **Modules** – Ensure Microsoft.Graph is imported (if using modules)
- 🌐 **Connectivity** – Validate access to Microsoft Graph and SharePoint
- 🔐 **Permissions** – Confirm all Graph and SharePoint permissions are granted
- ⚙️ **Resources** – Check runbook limits, job quotas, and timeout settings
- 📄 **Logs** – Review detailed runbook logs in Azure for stack traces and errors

---

## 👤 Author
Although this script has been tested extensively and successfully, PLEASE ensure you are thoroughly testing this on your own prior to introducing this into production.
Use at your own risk. 


**Eddie Jimenez**  
@edtrax


