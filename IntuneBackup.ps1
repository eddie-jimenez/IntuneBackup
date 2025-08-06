<#
.SYNOPSIS
    Comprehensive Intune Backup Solution using Direct Microsoft Graph API
    
.DESCRIPTION
    This Azure Automation runbook provides complete Intune configuration backup functionality
    using direct Microsoft Graph API calls with Managed Identity authentication. Originally 
    developed to replace the IntuneBackupAndRestore PowerShell module which doesn't support 
    Managed Identity in Azure Automation environments.
    
    Key Features:
    - Comprehensive backup coverage matching IntuneBackupAndRestore module capabilities
    - Individual JSON files for each configuration (not consolidated files)
    - Organized folder structure with platform-specific subfolders for applications  
    - Direct SharePoint upload with automatic year/month folder creation
    - Professional HTML email notifications (green success, red failure themes)
    - Accurate runtime duration tracking including SharePoint upload time
    - Robust error handling with detailed troubleshooting information
    - PowerShell 5.1 compatible with Azure Automation runtime constraints
    - Rate limiting and retry logic for Graph API calls
    
    Backup Categories Include:
    - Device Configurations & Compliance Policies
    - Settings Catalog & Configuration Policies  
    - Device Management Scripts & Health Scripts
    - Administrative Templates & ADMX Files
    - Mobile Applications (organized by platform)
    - App Protection & Configuration Policies
    - Autopilot Deployment Profiles & Enrollment Settings
    - Windows Update Profiles (Feature, Quality, Driver)
    - Conditional Access Policies & Named Locations
    - Security Baselines & Device Management Intents
    - Role Definitions & Assignment Filters
    - Mac Scripts & Custom Attributes
    - And 20+ additional Intune configuration categories
    
.PARAMETER EmailRecipient
    Comma-separated list of email addresses to receive backup status notifications.
    Default: 'YOUR TEST EMAIL
    
.EXAMPLE
    # Run with default email recipient
    .\IntuneBackup.ps1
    
.EXAMPLE  
    # Run with multiple email recipients
    .\IntuneBackup.ps1 -EmailRecipient 'admin1@company.com,admin2@company.com'
    
.NOTES
    Prerequisites:
    - Azure Automation Account with System-assigned Managed Identity enabled
    - Required Microsoft Graph API permissions on Managed Identity:
      * DeviceManagementConfiguration.Read.All
      * DeviceManagementApps.Read.All  
      * DeviceManagementRBAC.Read.All
      * DeviceManagementManagedDevices.Read.All
      * Mail.Send
      * Sites.ReadWrite.All
      * Directory.Read.All
      * Policy.Read.All (for Conditional Access)
    - Required Azure Automation Modules:
      * Microsoft.Graph.Authentication
      * Microsoft.Graph.Mail
    - SharePoint site permissions for backup storage
    
    SharePoint Structure:
    /sites/YOUR TEAM SITE/Intune Backups/YYYY/MM/
    
    
.AUTHOR
    Eddie Jimenez @edtrax
    
.VERSION
    6.0 - Complete Production Solution
    
.LASTMODIFIED
    August 2025
    
.CHANGELOG
    v6.0 - Complete production solution with full category coverage and professional styling
    v5.2 - Fixed duration calculation to include SharePoint upload time
    v5.1 - Enhanced email formatting and individual file generation
    v5.0 - Direct Graph API implementation to replace module dependency
    v4.0 - Authentication fixes for Managed Identity compatibility  
    v3.0 - PowerShell 5.1 runtime compatibility improvements
    v2.0 - SharePoint integration and comprehensive error handling
    v1.0 - Initial IntuneBackupAndRestore module approach (deprecated)
    
.LINK
    https://github.com/jseerden/IntuneBackupAndRestore (original module inspiration)
    https://docs.microsoft.com/en-us/graph/api/overview (Microsoft Graph API)
    https://docs.microsoft.com/en-us/azure/automation/ (Azure Automation documentation)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$EmailRecipient = 'YOUR RECIPIENT EMAIL' # Update with your recipient email. This can be a testing email. This can also be adjusted within the runbook settings or schedules. 
)

#Store Start Time
$scriptStart = Get-Date

# Environment check
if (-not $PSPrivateMetadata.JobId.Guid) {
    Write-Error "This script requires Azure Automation"
    exit 1
}

Write-Output "Running inside Azure Automation Runbook"

# Import required modules
Write-Output "=== Module Setup ==="
try {
    Import-Module Microsoft.Graph.Authentication -Force
    Import-Module Microsoft.Graph.Mail -Force
    Write-Output "✅ Modules imported"
} catch {
    Write-Error "Module import failed: $($_.Exception.Message)"
    exit 1
}

# Get authentication token
Write-Output "=== Authentication ==="
try {
    $resourceURI = "https://graph.microsoft.com"
    $tokenAuthURI = $env:IDENTITY_ENDPOINT + "?resource=$resourceURI&api-version=2019-08-01"
    $tokenResponse = Invoke-RestMethod -Method Get -Headers @{"X-IDENTITY-HEADER" = $env:IDENTITY_HEADER} -Uri $tokenAuthURI
    $accessToken = $tokenResponse.access_token
    
    $headers = @{
        "Authorization" = "Bearer $accessToken"
        "Content-Type" = "application/json"
    }
    
    Write-Output "✅ Authentication successful"
    
    # Test API
    $testResult = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations?`$top=1" -Headers $headers -Method GET
    Write-Output "✅ Graph API test passed"
    
} catch {
    Write-Error "Authentication failed: $($_.Exception.Message)"
    exit 1
}

# Function to get all pages
function Get-AllPages {
    param([string]$Uri)
    
    $results = New-Object System.Collections.ArrayList
    $nextUri = $Uri
    
    do {
        try {
            $response = Invoke-RestMethod -Uri $nextUri -Headers $headers -Method GET
            
            if ($response.value) {
                foreach ($item in $response.value) {
                    $null = $results.Add($item)
                }
            }
            
            $nextUri = $response.'@odata.nextLink'
            
            if ($nextUri) {
                Start-Sleep -Milliseconds 200
            }
            
        } catch {
            Write-Warning "Error getting $nextUri : $($_.Exception.Message)"
            break
        }
    } while ($nextUri)
    
    return $results
}

function Get-AppPlatformFolder {
    param($odataType)

    switch -Wildcard ($odataType) {
        "*win32LobApp"                  { return "Windows" }
        "*windowsStoreApp"              { return "Windows" }
        "*windowsMicrosoftEdgeApp"      { return "Windows" }
        "*macOS*App"                    { return "macOS" }
        "*ios*App"                      { return "iOS" }
        "*android*App"                  { return "Android" }
        "*webApp"                       { return "Web" }
        default                         { return "Other" }
    }
}

# Main backup execution
try {
    Write-Output "=== Starting Backup ==="
    
    $timestamp = Get-Date -Format "MM-dd-yyyy_HHmmss"
    $backupFolder = "$env:TEMP\IntuneBackup_$timestamp"
    $archiveName = "IntuneBackup_$timestamp.zip"
    
    New-Item -Path $backupFolder -ItemType Directory -Force | Out-Null
    Write-Output "Backup folder: $backupFolder"
    
    # Define backup categories. This array can be adjusted to add or comment out categories. 
    $categories = @{
        # Device Management
        "DeviceConfigurations" = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"
        "CompliancePolicies" = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies"
        "ConfigurationPolicies" = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$expand=settings"
        "DeviceManagementScripts" = "https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts"
        "DeviceHealthScripts" = "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts"
        "DeviceManagementIntents" = "https://graph.microsoft.com/beta/deviceManagement/intents"
        
        # Administrative Templates
        "AdministrativeTemplates" = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations"
        "ADMXFiles" = "https://graph.microsoft.com/beta/deviceManagement/groupPolicyUploadedDefinitionFiles"
        
        # Applications
        "MobileApps" = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$filter=(appAvailability eq null or appAvailability eq 'lineOfBusiness' or isAssigned eq true)&$orderby=displayName"
        "ManagedAppPolicies" = "https://graph.microsoft.com/beta/deviceAppManagement/managedAppPolicies"
        "AppConfigurationPolicies" = "https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations"
        "PolicySets" = "https://graph.microsoft.com/beta/deviceAppManagement/policySets"
        
        # Enrollment & Autopilot
        "WindowsAutopilotDeploymentProfiles" = "https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles"
        "EnrollmentConfigurations" = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations"
        "EnrollmentStatusPageProfiles" = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations?`$filter=deviceEnrollmentConfigurationType eq 'windows10EnrollmentCompletionPageConfiguration'"
        "EnrollmentRestrictions" = "https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations?`$filter=deviceEnrollmentConfigurationType eq 'limit' or deviceEnrollmentConfigurationType eq 'platformRestrictions'"
        
        # Settings Catalog (with settings expanded)
        "SettingsCatalog" = "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$expand=settings"

        # Updates
        "WindowsFeatureUpdateProfiles" = "https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles"
        "WindowsQualityUpdateProfiles" = "https://graph.microsoft.com/beta/deviceManagement/windowsQualityUpdateProfiles"
        "WindowsDriverUpdateProfiles" = "https://graph.microsoft.com/beta/deviceManagement/windowsDriverUpdateProfiles"
        "MacOSSoftwareUpdateAccountSummaries" = "https://graph.microsoft.com/beta/deviceManagement/macOSSoftwareUpdateAccountSummaries"
        
        # Security & Compliance
        "ConditionalAccessPolicies" = "https://graph.microsoft.com/beta/identity/conditionalAccess/policies"
        "NamedLocations" = "https://graph.microsoft.com/beta/identity/conditionalAccess/namedLocations"
        "AuthenticationStrengthPolicies" = "https://graph.microsoft.com/beta/identity/conditionalAccess/authenticationStrengths/policies"
        
        # Notifications & Branding
        "NotificationMessageTemplates" = "https://graph.microsoft.com/beta/deviceManagement/notificationMessageTemplates"
        "IntuneBrandingProfiles" = "https://graph.microsoft.com/beta/deviceManagement/intuneBrandingProfiles"
        "OrganizationalMessages" = "https://graph.microsoft.com/beta/deviceManagement/organizationalMessageDetails"
        
        # Filters & Assignment
        "AssignmentFilters" = "https://graph.microsoft.com/beta/deviceManagement/assignmentFilters"
        "RoleScopeTaggings" = "https://graph.microsoft.com/beta/deviceManagement/roleScopeTags"
        "RoleDefinitions" = "https://graph.microsoft.com/beta/deviceManagement/roleDefinitions"
        
        # Mac Management
        "MacScripts" = "https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts"
        "MacCustomAttributes" = "https://graph.microsoft.com/beta/deviceManagement/deviceCustomAttributeShellScripts"
        
        # Remote Actions & Settings
        #"RemoteActionAudits" = "https://graph.microsoft.com/beta/deviceManagement/remoteActionAudits?`$top=100"
        "DeviceCategories" = "https://graph.microsoft.com/beta/deviceManagement/deviceCategories"
        #"TermsAndConditions" = "https://graph.microsoft.com/beta/deviceManagement/termsAndConditions"
        
        # Partner Integrations
        #"DeviceCompliancePartners" = "https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePartners"
        "MobileThreatDefenseConnectors" = "https://graph.microsoft.com/beta/deviceManagement/mobileThreatDefenseConnectors"
        
        # Reusable Settings
        "ReusablePolicySettings" = "https://graph.microsoft.com/beta/deviceManagement/reusablePolicySettings"
        
        # Azure AD / Groups - NOTE: Depending on the size of your tenant, this could be MASSIVE. Recommend to leave this commented out unless needed. 
        #"Groups" = "https://graph.microsoft.com/beta/groups"
        #"DirectoryRoles" = "https://graph.microsoft.com/beta/directoryRoles"
        
        # Tenant Settings
        "DeviceManagementSettings" = "https://graph.microsoft.com/beta/deviceManagement/settings"
        "Organization" = "https://graph.microsoft.com/beta/organization"
    }
    
    $totalItems = 0
    $results = @{}
    
foreach ($categoryName in $categories.Keys) {
    try {
        Write-Output "Backing up $categoryName..."
        
        $items = Get-AllPages -Uri $categories[$categoryName]
        $itemCount = $items.Count
        
        if ($itemCount -gt 0) {
            $categoryPath = Join-Path $backupFolder $categoryName
            New-Item -Path $categoryPath -ItemType Directory -Force | Out-Null

            if ($categoryName -eq "MobileApps") {
                foreach ($item in $items) {
                    try {
                        # Determine platform from @odata.type
                        $odataType = $item.'@odata.type'
                        switch -Wildcard ($odataType) {
                            "*win32LobApp"                  { $platform = "Windows" }
                            "*windowsStoreApp"              { $platform = "Windows" }
                            "*windowsMicrosoftEdgeApp"      { $platform = "Windows" }
                            "*macOS*App"                    { $platform = "macOS" }
                            "*ios*App"                      { $platform = "iOS" }
                            "*android*App"                  { $platform = "Android" }
                            "*webApp"                       { $platform = "Web" }
                            default                         { $platform = "Other" }
                        }

                        $platformPath = Join-Path $categoryPath $platform
                        if (!(Test-Path $platformPath)) {
                            New-Item -Path $platformPath -ItemType Directory -Force | Out-Null
                        }

                        # Determine filename
                        $fileName = "Unknown"
                        if ($item.displayName) {
                            $fileName = $item.displayName
                        } elseif ($item.name) {
                            $fileName = $item.name
                        } elseif ($item.id) {
                            $fileName = $item.id
                        }

                        $cleanFileName = $fileName -replace '[\\/*?:"<>|]', '_'
                        $cleanFileName = $cleanFileName -replace '\s+', ' '
                        $cleanFileName = $cleanFileName.Trim()
                        if ($cleanFileName.Length -gt 100) {
                            $cleanFileName = $cleanFileName.Substring(0, 100)
                        }

                        $filePath = Join-Path $platformPath "$cleanFileName.json"
                        $counter = 1
                        while (Test-Path $filePath) {
                            $filePath = Join-Path $platformPath "$cleanFileName ($counter).json"
                            $counter++
                        }

                        $item | ConvertTo-Json -Depth 20 | Out-File $filePath -Encoding UTF8
                    } catch {
                        Write-Warning "Failed to save item in MobileApps : $($_.Exception.Message)"
                    }
                }

                # Save master list
                $items | ConvertTo-Json -Depth 20 | Out-File "$categoryPath\All_MobileApps.json" -Encoding UTF8

            } else {
                # Default handling for all other categories
                foreach ($item in $items) {
                    try {
                        $fileName = "Unknown"
                        if ($item.displayName) {
                            $fileName = $item.displayName
                        } elseif ($item.name) {
                            $fileName = $item.name
                        } elseif ($item.id) {
                            $fileName = $item.id
                        }

                        $cleanFileName = $fileName -replace '[\\/*?:"<>|]', '_'
                        $cleanFileName = $cleanFileName -replace '\s+', ' '
                        $cleanFileName = $cleanFileName.Trim()
                        if ($cleanFileName.Length -gt 100) {
                            $cleanFileName = $cleanFileName.Substring(0, 100)
                        }

                        $filePath = Join-Path $categoryPath "$cleanFileName.json"
                        $counter = 1
                        while (Test-Path $filePath) {
                            $filePath = Join-Path $categoryPath "$cleanFileName ($counter).json"
                            $counter++
                        }

                        $item | ConvertTo-Json -Depth 20 | Out-File $filePath -Encoding UTF8
                    } catch {
                        Write-Warning "Failed to save item in $categoryName : $($_.Exception.Message)"
                    }
                }

                # Save master list
                $items | ConvertTo-Json -Depth 20 | Out-File "$categoryPath\All_$categoryName.json" -Encoding UTF8
            }

            Write-Output "✅ $categoryName : $itemCount items"
            $results[$categoryName] = $itemCount
            $totalItems += $itemCount

        } else {
            Write-Output "⚠️ $categoryName : No items found"
            $results[$categoryName] = 0
        }

    } catch {
        Write-Warning "$categoryName failed: $($_.Exception.Message)"
        $results[$categoryName] = 0
    }
}

    Write-Output "✅ Backup completed: $totalItems total items"
    
         # Create summary
    $summary = @{
        BackupDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        TotalItems = $totalItems
        Categories = $results
        Method = "Direct Graph API"
        Duration = $durationFormatted
        SharePointUpload = $sharePointSuccess
    }

    $summaryPath = Join-Path $backupFolder "Summary.json"
    $summary | ConvertTo-Json -Depth 5 | Out-File $summaryPath -Encoding UTF8   

    # Verify files were created
    $backupFiles = Get-ChildItem -Path $backupFolder -Recurse -File
    if ($backupFiles.Count -eq 0) {
        throw "No backup files were created"
    }

    Write-Output "Files created: $($backupFiles.Count)"
    Write-Output "Total size: $([math]::Round(($backupFiles | Measure-Object Length -Sum).Sum / 1MB, 2)) MB"



    # SharePoint upload
    Write-Output "Uploading extracted backup to SharePoint..."
    try {
        $year = (Get-Date).ToString("yyyy")
        $month = (Get-Date).ToString("MM")
        
        # Get SharePoint site
        $siteUri = "YOUR SITE URI" # Update with your site URI. Example: `https://graph.microsoft.com/v1.0/sites/{tenant}.sharepoint.com:/sites/{site-name}
        $site = Invoke-RestMethod -Uri $siteUri -Headers $headers -Method GET
        $siteId = $site.id

        # Get document library
        $driveUri = "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
        $drives = Invoke-RestMethod -Uri $driveUri -Headers $headers -Method GET
        $documentsLibrary = $drives.value | Where-Object { $_.name -eq "Documents" }
        $driveId = $documentsLibrary.id

        # Create year/month folder structure
        $folderSegments = @("Intune Backups", $year, $month)
        $currentPath = ""
        foreach ($segment in $folderSegments) {
            $currentPath = if ($currentPath) { "$currentPath/$segment" } else { $segment }
            $checkUri = "https://graph.microsoft.com/v1.0/drives/$driveId/root:/$currentPath"
            try {
                Invoke-RestMethod -Uri $checkUri -Headers $headers -Method GET | Out-Null
            } catch {
                $createUri = "https://graph.microsoft.com/v1.0/drives/$driveId/root:/$($currentPath.Substring(0, $currentPath.LastIndexOf('/'))):/children"
                $newFolderBody = @{ name = $segment; folder = @{} } | ConvertTo-Json
                Invoke-RestMethod -Uri $createUri -Headers $headers -Method POST -Body $newFolderBody | Out-Null
            }
        }

        # Upload each file
        foreach ($file in $backupFiles) {
            $relativePath = $file.FullName.Substring($backupFolder.Length).TrimStart('\', '/')
            $sharePointPath = "$currentPath/$relativePath" -replace '\\', '/'

            $uploadUri = "https://graph.microsoft.com/v1.0/drives/$driveId/root:/${sharePointPath}:/content"
            try {
                $fileBytes = [System.IO.File]::ReadAllBytes($file.FullName)
                Invoke-RestMethod -Uri $uploadUri -Headers $headers -Method PUT -Body $fileBytes | Out-Null
                Write-Output "✅ Uploaded: $sharePointPath"
            } catch {
                Write-Warning "❌ Failed to upload $relativePath : $($_.Exception.Message)"
            }
        }

        $sharePointSuccess = $true
        $sharePointPath = "/sites/YOUR SHAREPOINT PATH/$currentPath"   # Update with your Sharepoint Online path. Example: ITadmins/Intune Backups

        # Construct a manual SharePoint web URL for the uploaded folder
        Add-Type -AssemblyName System.Web
        $encodedPath = [System.Web.HttpUtility]::UrlEncode($sharePointPath)
        $shareableLink = "YOUR SHAREPOINT URL$encodedPath"   # Update with your Sharepoint Online URL. Example: https://{tenant}.sharepoint.com/sites/{site-name}/

    } catch {
        Write-Warning "SharePoint upload failed: $($_.Exception.Message)"
        $sharePointSuccess = $false
        $sharePointPath = "Upload failed"
    }

        # Calculate Script Runtime after operations complete
        $scriptEnd = Get-Date
        $duration = $scriptEnd - $scriptStart
        $durationFormatted = "{0:hh\:mm\:ss}" -f $duration

    
    # Send email
    try {
        Connect-MgGraph -AccessToken (ConvertTo-SecureString -String $accessToken -AsPlainText -Force) -NoWelcome
        
        $emailBody = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; background-color: #f5f5f5; }
        .container { max-width: 600px; margin: 20px auto; background-color: white; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .header { background: linear-gradient(135deg, #28a745, #20c997); color: white; padding: 30px 20px; text-align: center; }
        .header h1 { margin: 0; font-size: 28px; font-weight: 600; }
        .header p { margin: 10px 0 0 0; font-size: 16px; opacity: 0.9; }
        .content { padding: 30px 20px; }
        .status-badge { display: inline-block; background-color: #28a745; color: white; padding: 8px 16px; border-radius: 20px; font-weight: 600; font-size: 14px; margin-bottom: 20px; }
        .metric { display: flex; justify-content: space-between; align-items: center; padding: 12px 0; border-bottom: 1px solid #eee; }
        .metric:last-child { border-bottom: none; }
        .metric-label { font-weight: 600; color: #333; }
        .metric-value { color: #666; font-family: 'Courier New', monospace; }
        .sharepoint-link { display: inline-block; background-color: #0078d4; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; margin: 20px 0; font-weight: 600; }
        .sharepoint-link:hover { background-color: #106ebe; }
        .backup-section { margin-top: 30px; }
        .backup-section h3 { color: #333; margin-bottom: 15px; border-bottom: 2px solid #28a745; padding-bottom: 5px; }
        .backup-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
        .backup-item { background-color: #f8f9fa; padding: 8px 12px; border-radius: 4px; font-size: 14px; }
        .backup-item strong { color: #28a745; }
        .footer { background-color: #f8f9fa; padding: 20px; text-align: center; color: #666; font-size: 12px; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🔄 Intune Monthly Backup Report</h1>
        </div>
        
        <div class="content">
            <div class="status-badge">✅ BACKUP SUCCESSFUL</div>
            
            <div class="metric">
                <span class="metric-label">Backup Time</span>
                <span class="metric-value">$timestamp</span>
            </div>
            <div class="metric">
                <span class="metric-label">Duration</span>
                <span class="metric-value">$durationFormatted</span>
            </div>
            <div class="metric">
                <span class="metric-label">Total Items</span>
                <span class="metric-value">$totalItems configurations</span>
            </div>
            
            $(if ($sharePointSuccess) {
                "<div class='metric'><span class='metric-label'>SharePoint</span><span class='metric-value'>✅ Successfully uploaded</span></div><a href='$shareableLink' class='sharepoint-link'>📂 View Backup in SharePoint</a>"
            } else {
                "<div class='metric'><span class='metric-label'>SharePoint</span><span class='metric-value'>⚠️ Upload failed</span></div>"
            })

            <div class="backup-section">
                <h3>📊 Backup Summary</h3>
                <div class="backup-grid">
                    $(foreach ($cat in $results.Keys | Sort-Object) {
                        if ($results[$cat] -gt 0) {
                            "<div class='backup-item'><strong>$($results[$cat])</strong> $cat</div>"
                        }
                    })
                </div>
            </div>
        </div>
        
        <div class="footer">
            <p><strong>Note:</strong> This custom in-house solution uses Managed Identity and direct MS Graph API calls to automate Intune backups.</p>
            <p><em>Generated automatically by IT Automation on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')</em></p>
        </div>
    </div>
</body>
</html>
"@
        
        $emailRecipients = $EmailRecipient -split ',' | ForEach-Object { $_.Trim() }
        
        foreach ($recipient in $emailRecipients) {
            $message = @{
                subject = "✅ Intune Backup Success - $timestamp"
                body = @{
                    contentType = "HTML"
                    content = $emailBody
                }
                toRecipients = @(
                    @{ emailAddress = @{ address = $recipient } }
                )
            }
            
            $requestBody = @{ message = $message } | ConvertTo-Json -Depth 10
            $emailUri = "https://graph.microsoft.com/v1.0/users/YOUR AUTOMATION ACCOUNT EMAIL/sendMail"      # Update with your automation account email. Example: automation@yourdomain.com
            Invoke-MgGraphRequest -Uri $emailUri -Method POST -Body $requestBody -ContentType "application/json"
        }
        
        Write-Output "✅ Success email sent"
        
    } catch {
        Write-Warning "Email failed: $($_.Exception.Message)"
    }
    
    # Cleanup
    Remove-Item -Path $backupFolder -Recurse -Force -ErrorAction SilentlyContinue
    
    # Final summary
    Write-Output ""
    Write-Output "BACKUP COMPLETE"
    Write-Output "=================="
    Write-Output "Status: SUCCESS"
    Write-Output "Time: $timestamp"
    Write-Output "Duration: $durationFormatted"
    Write-Output "Items: $totalItems"
    Write-Output "SharePoint: $(if ($sharePointSuccess) { 'Success' } else { 'Failed' })"
    Write-Output ""
    
} catch {
    # Calculate duration even in failure case
    $scriptEnd = Get-Date
    $duration = $scriptEnd - $scriptStart
    $durationFormatted = "{0:hh\:mm\:ss}" -f $duration
    
    Write-Error "❌ Backup failed after $durationFormatted : $($_.Exception.Message)"
    
    # Send failure email
    try {
        Connect-MgGraph -AccessToken (ConvertTo-SecureString -String $accessToken -AsPlainText -Force) -NoWelcome
        
        $failureBody = @"
<!DOCTYPE html>
<html>
<head>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; background-color: #f5f5f5; }
        .container { max-width: 600px; margin: 20px auto; background-color: white; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
        .header { background: linear-gradient(135deg, #dc3545, #c82333); color: white; padding: 30px 20px; text-align: center; }
        .header h1 { margin: 0; font-size: 28px; font-weight: 600; }
        .header p { margin: 10px 0 0 0; font-size: 16px; opacity: 0.9; }
        .content { padding: 30px 20px; }
        .status-badge { display: inline-block; background-color: #dc3545; color: white; padding: 8px 16px; border-radius: 20px; font-weight: 600; font-size: 14px; margin-bottom: 20px; }
        .metric { display: flex; justify-content: space-between; align-items: center; padding: 12px 0; border-bottom: 1px solid #eee; }
        .metric:last-child { border-bottom: none; }
        .metric-label { font-weight: 600; color: #333; }
        .metric-value { color: #666; font-family: 'Courier New', monospace; }
        .error-section { margin-top: 20px; background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 4px; padding: 15px; }
        .error-section h4 { margin-top: 0; color: #721c24; }
        .error-text { background-color: #fff; border: 1px solid #dee2e6; border-radius: 3px; padding: 10px; font-family: 'Courier New', monospace; font-size: 12px; color: #495057; white-space: pre-wrap; }
        .footer { background-color: #f8f9fa; padding: 20px; text-align: center; color: #666; font-size: 12px; }
        .troubleshooting { margin-top: 20px; background-color: #e2e3e5; border-radius: 4px; padding: 15px; }
        .troubleshooting h4 { margin-top: 0; color: #383d41; }
        .troubleshooting ul { margin: 10px 0; padding-left: 20px; }
        .troubleshooting li { margin: 5px 0; color: #495057; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>❌ Intune Monthly Backup FAILED</h1>
        </div>
        
        <div class="content">
            <div class="status-badge">❌ BACKUP FAILED</div>
            
            <div class="metric">
                <span class="metric-label">Failure Time</span>
                <span class="metric-value">$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</span>
            </div>
            <div class="metric">
                <span class="metric-label">Duration</span>
                <span class="metric-value">$durationFormatted</span>
            </div>

            <div class="error-section">
                <h4>💥 Error Details</h4>
                <div class="error-text">$($_.Exception.Message)</div>
            </div>

            <div class="troubleshooting">
                <h4>🔧 Troubleshooting Steps</h4>
                <ul>
                    <li>Check that the Managed Identity has the required Graph API permissions</li>
                    <li>Verify Microsoft.Graph modules are properly imported in Azure Automation</li>
                    <li>Review the Azure Automation logs for detailed error information</li>
                    <li>Ensure SharePoint site permissions are configured correctly</li>
                    <li>Check if there are any network connectivity issues</li>
                </ul>
            </div>
        </div>
        
        <div class="footer">
            <p><strong>Action Required:</strong> Please investigate the error above and re-run the backup manually if needed.</p>
            <p><em>Generated automatically by Azure Automation on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss UTC')</em></p>
        </div>
    </div>
</body>
</html>
"@
        
        $emailRecipients = $EmailRecipient -split ',' | ForEach-Object { $_.Trim() }
        
        foreach ($recipient in $emailRecipients) {
            $message = @{
                subject = "❌ Intune Backup FAILED - $(Get-Date -Format 'yyyy-MM-dd_HHmmss')"
                body = @{
                    contentType = "HTML"
                    content = $failureBody
                }
                toRecipients = @(
                    @{ emailAddress = @{ address = $recipient } }
                )
            }
            
            $requestBody = @{ message = $message } | ConvertTo-Json -Depth 10
            $emailUri = "https://graph.microsoft.com/v1.0/users/YOUR AUTOMATION ACCOUNT EMAIL/sendMail".      # Update with your automation account email. Example: automation@yourdomain.com
            Invoke-MgGraphRequest -Uri $emailUri -Method POST -Body $requestBody -ContentType "application/json"
        }
        
    } catch {
        Write-Error "Failed to send failure email: $($_.Exception.Message)"
    }
    
    exit 1
}
