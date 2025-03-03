# Microsoft 365 Administration Toolkit

A PowerShell script collection designed to simplify and automate common Microsoft 365 administration tasks. Perfect for IT administrators looking to streamline their workflows and demonstrate technical proficiency with Microsoft 365 services.

## Overview

This toolkit provides a set of PowerShell functions that help manage various aspects of Microsoft 365, including:

- User management (creation, license assignment, reporting)
- Microsoft Teams administration
- SharePoint Online permission management
- Usage reporting and analytics
- Security policies (Conditional Access)

## Prerequisites

To use this toolkit, you'll need:

- PowerShell 5.1 or higher
- Microsoft 365 administrator account
- The following PowerShell modules installed:
  - Microsoft.Online.SharePoint.PowerShell
  - ExchangeOnlineManagement
  - MicrosoftTeams
  - AzureAD
  - AzureADPreview (for Conditional Access policies)

## Installation

1. Clone this repository or download the script files
2. Install the required PowerShell modules:

```powershell
Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Force
Install-Module -Name ExchangeOnlineManagement -Force
Install-Module -Name MicrosoftTeams -Force
Install-Module -Name AzureAD -Force
Install-Module -Name AzureADPreview -Force
```

3. Import the toolkit in your PowerShell session:

```powershell
. .\M365AdminToolkit.ps1
```

## Key Features

### 1. Single Connection Function for All Microsoft 365 Services

The `Connect-M365Services` function provides a unified way to connect to multiple Microsoft 365 services with a single credential.

```powershell
$credential = Get-Credential
Connect-M365Services -Credential $credential -ExchangeOnline -SharePointOnline -MicrosoftTeams -AzureAD
```

### 2. Bulk User Management

Create multiple users at once from a CSV file using `New-M365BulkUsers`. The CSV should include FirstName, LastName, and Domain columns.

```powershell
New-M365BulkUsers -CsvPath "C:\Users.csv" -DefaultPassword "TempP@ss123!" -LicenseSkuId "tenant:ENTERPRISEPACK"
```

### 3. License Reporting

Generate comprehensive reports on user license assignments using `Export-M365UserLicenses`.

```powershell
Export-M365UserLicenses -OutputPath "C:\M365UserLicenses.csv" -IncludeDisabledUsers
```

### 4. Microsoft Teams Management

Create Teams with channels and add members in a single operation using `New-M365TeamWithChannels`.

```powershell
New-M365TeamWithChannels -TeamName "Project X" -TeamDescription "Team for Project X" -Channels @("General", "Planning", "Development") -OwnerEmails @("admin@contoso.com") -MemberEmails @("user1@contoso.com", "user2@contoso.com")
```

### 5. SharePoint Permission Management

Easily manage SharePoint site permissions with `Set-SPOSitePermissions`.

```powershell
Set-SPOSitePermissions -SiteUrl "https://contoso.sharepoint.com/sites/Marketing" -OwnersToAdd @{"John Doe" = "john@contoso.com"} -MembersToAdd @{"Jane Smith" = "jane@contoso.com"} -VisitorsToRemove @("formeruser@contoso.com")
```

### 6. Usage Reporting

Generate reports on Microsoft 365 service usage with `Get-M365UsageReport`.

```powershell
Get-M365UsageReport -OutputPath "C:\M365UsageReport.csv" -Days 30 -IncludeTeams -IncludeOneDrive -IncludeExchange -IncludeSharePoint
```

### 7. Security Policy Management

Create and manage Conditional Access policies with `New-M365ConditionalAccessPolicy`.

```powershell
New-M365ConditionalAccessPolicy -PolicyName "Require MFA for All Users" -RequireMFA
```

## Example Usage Scenarios

### Onboarding New Department

```powershell
# Connect to services
$credential = Get-Credential
Connect-M365Services -Credential $credential -ExchangeOnline -SharePointOnline -MicrosoftTeams -AzureAD

# Create users from CSV
New-M365BulkUsers -CsvPath "C:\MarketingTeam.csv" -DefaultPassword "Welcome2023!" -LicenseSkuId "tenant:ENTERPRISEPACK"

# Create Teams workspace
New-
