# Microsoft 365 Administration Toolkit
# A collection of PowerShell scripts to automate common Microsoft 365 administration tasks

# Function to connect to Microsoft 365 services
function Connect-M365Services {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [PSCredential]$Credential,
        
        [Parameter()]
        [switch]$ExchangeOnline,
        
        [Parameter()]
        [switch]$SharePointOnline,
        
        [Parameter()]
        [switch]$MicrosoftTeams,
        
        [Parameter()]
        [switch]$AzureAD
    )
    
    try {
        # Connect to Exchange Online if requested
        if ($ExchangeOnline) {
            Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
            Connect-ExchangeOnline -Credential $Credential -ErrorAction Stop
            Write-Host "Successfully connected to Exchange Online." -ForegroundColor Green
        }
        
        # Connect to SharePoint Online if requested
        if ($SharePointOnline) {
            Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
            $orgName = ($Credential.UserName -split '@')[1] -replace '\..*$', ''
            Connect-SPOService -Url "https://$orgName-admin.sharepoint.com" -Credential $Credential -ErrorAction Stop
            Write-Host "Successfully connected to SharePoint Online." -ForegroundColor Green
        }
        
        # Connect to Microsoft Teams if requested
        if ($MicrosoftTeams) {
            Write-Host "Connecting to Microsoft Teams..." -ForegroundColor Cyan
            Connect-MicrosoftTeams -Credential $Credential -ErrorAction Stop
            Write-Host "Successfully connected to Microsoft Teams." -ForegroundColor Green
        }
        
        # Connect to Azure AD if requested
        if ($AzureAD) {
            Write-Host "Connecting to Azure AD..." -ForegroundColor Cyan
            Connect-AzureAD -Credential $Credential -ErrorAction Stop
            Write-Host "Successfully connected to Azure AD." -ForegroundColor Green
        }
        
        return $true
    }
    catch {
        Write-Error "Error connecting to Microsoft 365 services: $_"
        return $false
    }
}

# Function to create new users in bulk from CSV
function New-M365BulkUsers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$CsvPath,
        
        [Parameter(Mandatory = $true)]
        [string]$DefaultPassword,
        
        [Parameter()]
        [string]$UsageLocation = "US",
        
        [Parameter()]
        [string[]]$LicenseSkuId
    )
    
    try {
        # Check if connected to Azure AD
        try {
            Get-AzureADTenantDetail -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Error "Not connected to Azure AD. Please run Connect-M365Services -AzureAD first."
            return
        }
        
        # Import CSV file
        if (-not (Test-Path $CsvPath)) {
            Write-Error "CSV file not found: $CsvPath"
            return
        }
        
        $users = Import-Csv -Path $CsvPath
        $securePassword = ConvertTo-SecureString -String $DefaultPassword -AsPlainText -Force
        
        $results = @()
        
        foreach ($user in $users) {
            try {
                Write-Host "Processing user: $($user.FirstName) $($user.LastName)" -ForegroundColor Cyan
                
                # Create user principal name
                $upn = "$($user.FirstName).$($user.LastName)@$($user.Domain)"
                
                # Create new user
                $newUser = New-AzureADUser -DisplayName "$($user.FirstName) $($user.LastName)" `
                                          -GivenName $user.FirstName `
                                          -Surname $user.LastName `
                                          -UserPrincipalName $upn `
                                          -MailNickName "$($user.FirstName)$($user.LastName)" `
                                          -UsageLocation $UsageLocation `
                                          -AccountEnabled $true `
                                          -PasswordProfile @{
                                              Password = $DefaultPassword
                                              ForceChangePasswordNextLogin = $true
                                          }
                
                # Assign license if specified
                if ($LicenseSkuId) {
                    foreach ($skuId in $LicenseSkuId) {
                        Set-AzureADUserLicense -ObjectId $newUser.ObjectId -AssignedLicenses @{
                            SkuId = $skuId
                        }
                    }
                }
                
                $results += [PSCustomObject]@{
                    UserPrincipalName = $upn
                    Status = "Created"
                    ErrorDetails = ""
                }
                
                Write-Host "Successfully created user: $upn" -ForegroundColor Green
            }
            catch {
                $results += [PSCustomObject]@{
                    UserPrincipalName = $upn
                    Status = "Failed"
                    ErrorDetails = $_.Exception.Message
                }
                
                Write-Host "Failed to create user: $upn - $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        
        return $results
    }
    catch {
        Write-Error "Error in bulk user creation: $_"
    }
}

# Function to export all Microsoft 365 users with license information
function Export-M365UserLicenses {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter()]
        [switch]$IncludeDisabledUsers
    )
    
    try {
        # Check if connected to Azure AD
        try {
            Get-AzureADTenantDetail -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Error "Not connected to Azure AD. Please run Connect-M365Services -AzureAD first."
            return
        }
        
        # Get all subscribed SKUs (licenses)
        $licenses = Get-AzureADSubscribedSku
        $licenseTable = @{}
        foreach ($license in $licenses) {
            $licenseTable[$license.SkuId] = $license.SkuPartNumber
        }
        
        # Get users
        Write-Host "Retrieving users from Azure AD..." -ForegroundColor Cyan
        $filter = if (-not $IncludeDisabledUsers) { "AccountEnabled eq true" } else { $null }
        $users = Get-AzureADUser -All $true -Filter $filter
        
        $results = @()
        $count = 0
        $total = $users.Count
        
        foreach ($user in $users) {
            $count++
            Write-Progress -Activity "Processing user licenses" -Status "$count of $total" -PercentComplete (($count / $total) * 100)
            
            $userLicenses = @()
            foreach ($license in $user.AssignedLicenses) {
                $licenseName = $licenseTable[$license.SkuId]
                $userLicenses += $licenseName
            }
            
            $results += [PSCustomObject]@{
                DisplayName = $user.DisplayName
                UserPrincipalName = $user.UserPrincipalName
                Enabled = $user.AccountEnabled
                Department = $user.Department
                JobTitle = $user.JobTitle
                Licenses = ($userLicenses -join ", ")
                LicenseCount = $userLicenses.Count
            }
        }
        
        Write-Progress -Activity "Processing user licenses" -Completed
        
        # Export to CSV
        $results | Export-Csv -Path $OutputPath -NoTypeInformation
        Write-Host "Successfully exported $($results.Count) users to $OutputPath" -ForegroundColor Green
        
        return $results
    }
    catch {
        Write-Error "Error exporting user licenses: $_"
    }
}

# Function to manage Microsoft Teams
function New-M365TeamWithChannels {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$TeamName,
        
        [Parameter(Mandatory = $true)]
        [string]$TeamDescription,
        
        [Parameter()]
        [string[]]$Channels,
        
        [Parameter()]
        [string[]]$OwnerEmails,
        
        [Parameter()]
        [string[]]$MemberEmails
    )
    
    try {
        # Check if connected to Microsoft Teams
        try {
            Get-Team -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Error "Not connected to Microsoft Teams. Please run Connect-M365Services -MicrosoftTeams first."
            return
        }
        
        # Create new team
        Write-Host "Creating new team: $TeamName" -ForegroundColor Cyan
        $team = New-Team -DisplayName $TeamName -Description $TeamDescription -Visibility "Private"
        
        # Add owners if specified
        if ($OwnerEmails) {
            foreach ($owner in $OwnerEmails) {
                Add-TeamUser -GroupId $team.GroupId -User $owner -Role "Owner"
                Write-Host "Added owner to team: $owner" -ForegroundColor Green
            }
        }
        
        # Add members if specified
        if ($MemberEmails) {
            foreach ($member in $MemberEmails) {
                Add-TeamUser -GroupId $team.GroupId -User $member -Role "Member"
                Write-Host "Added member to team: $member" -ForegroundColor Green
            }
        }
        
        # Create channels if specified
        if ($Channels) {
            foreach ($channel in $Channels) {
                New-TeamChannel -GroupId $team.GroupId -DisplayName $channel
                Write-Host "Created channel: $channel" -ForegroundColor Green
            }
        }
        
        Write-Host "Successfully created team: $TeamName with ID: $($team.GroupId)" -ForegroundColor Green
        return $team
    }
    catch {
        Write-Error "Error creating team: $_"
    }
}

# Function to manage SharePoint site permissions
function Set-SPOSitePermissions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter()]
        [hashtable]$OwnersToAdd,
        
        [Parameter()]
        [hashtable]$MembersToAdd,
        
        [Parameter()]
        [hashtable]$VisitorsToAdd,
        
        [Parameter()]
        [string[]]$OwnersToRemove,
        
        [Parameter()]
        [string[]]$MembersToRemove,
        
        [Parameter()]
        [string[]]$VisitorsToRemove,
        
        [Parameter()]
        [switch]$DisableSharing,
        
        [Parameter()]
        [switch]$EnableExternalSharing
    )
    
    try {
        # Check if connected to SharePoint Online
        try {
            Get-SPOTenant -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Error "Not connected to SharePoint Online. Please run Connect-M365Services -SharePointOnline first."
            return
        }
        
        # Get site object
        $site = Get-SPOSite -Identity $SiteUrl -ErrorAction Stop
        
        # Add owners
        if ($OwnersToAdd) {
            foreach ($owner in $OwnersToAdd.Keys) {
                $email = $OwnersToAdd[$owner]
                Set-SPOUser -Site $SiteUrl -LoginName $email -IsSiteCollectionAdmin $true
                Write-Host "Added owner: $owner ($email)" -ForegroundColor Green
            }
        }
        
        # Add members
        if ($MembersToAdd) {
            foreach ($member in $MembersToAdd.Keys) {
                $email = $MembersToAdd[$member]
                Add-SPOUser -Site $SiteUrl -LoginName $email -Group "$($site.Title) Members"
                Write-Host "Added member: $member ($email)" -ForegroundColor Green
            }
        }
        
        # Add visitors
        if ($VisitorsToAdd) {
            foreach ($visitor in $VisitorsToAdd.Keys) {
                $email = $VisitorsToAdd[$visitor]
                Add-SPOUser -Site $SiteUrl -LoginName $email -Group "$($site.Title) Visitors"
                Write-Host "Added visitor: $visitor ($email)" -ForegroundColor Green
            }
        }
        
        # Remove owners
        if ($OwnersToRemove) {
            foreach ($email in $OwnersToRemove) {
                Set-SPOUser -Site $SiteUrl -LoginName $email -IsSiteCollectionAdmin $false
                Write-Host "Removed owner: $email" -ForegroundColor Yellow
            }
        }
        
        # Remove members and visitors (requires PnP)
        if (($MembersToRemove -or $VisitorsToRemove) -and (Get-Command Connect-PnPOnline -ErrorAction SilentlyContinue)) {
            # Connect to PnP
            $credential = Get-Credential -Message "Enter credentials for PnP connection"
            Connect-PnPOnline -Url $SiteUrl -Credentials $credential
            
            if ($MembersToRemove) {
                foreach ($email in $MembersToRemove) {
                    Remove-PnPGroupMember -LoginName $email -Group "$($site.Title) Members"
                    Write-Host "Removed member: $email" -ForegroundColor Yellow
                }
            }
            
            if ($VisitorsToRemove) {
                foreach ($email in $VisitorsToRemove) {
                    Remove-PnPGroupMember -LoginName $email -Group "$($site.Title) Visitors"
                    Write-Host "Removed visitor: $email" -ForegroundColor Yellow
                }
            }
        }
        
        # Configure sharing settings
        if ($DisableSharing) {
            Set-SPOSite -Identity $SiteUrl -SharingCapability Disabled
            Write-Host "Disabled sharing for site: $SiteUrl" -ForegroundColor Yellow
        }
        
        if ($EnableExternalSharing) {
            Set-SPOSite -Identity $SiteUrl -SharingCapability ExternalUserSharingOnly
            Write-Host "Enabled external sharing for site: $SiteUrl" -ForegroundColor Green
        }
        
        Write-Host "Successfully updated permissions for site: $SiteUrl" -ForegroundColor Green
    }
    catch {
        Write-Error "Error managing site permissions: $_"
    }
}

# Function to generate Microsoft 365 usage report
function Get-M365UsageReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$OutputPath,
        
        [Parameter()]
        [ValidateSet("7", "30", "90", "180")]
        [string]$Days = "30",
        
        [Parameter()]
        [switch]$IncludeTeams,
        
        [Parameter()]
        [switch]$IncludeOneDrive,
        
        [Parameter()]
        [switch]$IncludeExchange,
        
        [Parameter()]
        [switch]$IncludeSharePoint
    )
    
    try {
        # Check if connected to Exchange Online (required for reports)
        try {
            Get-OrganizationConfig -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Error "Not connected to Exchange Online. Please run Connect-M365Services -ExchangeOnline first."
            return
        }
        
        $reports = @()
        
        # Get Teams user activity
        if ($IncludeTeams) {
            Write-Host "Retrieving Microsoft Teams usage data..." -ForegroundColor Cyan
            $teamsReport = Get-TeamsUserActivityUserDetail -Period $Days
            $reports += [PSCustomObject]@{
                Service = "Microsoft Teams"
                ActiveUsers = ($teamsReport | Where-Object { $_.TeamsChannelMessages -gt 0 -or $_.TeamsChatMessages -gt 0 }).Count
                TotalUsers = $teamsReport.Count
                UsagePercentage = [math]::Round((($teamsReport | Where-Object { $_.TeamsChannelMessages -gt 0 -or $_.TeamsChatMessages -gt 0 }).Count / $teamsReport.Count) * 100, 2)
            }
        }
        
        # Get OneDrive user activity
        if ($IncludeOneDrive) {
            Write-Host "Retrieving OneDrive for Business usage data..." -ForegroundColor Cyan
            $oneDriveReport = Get-OneDriveUsageAccountDetail -Period $Days
            $reports += [PSCustomObject]@{
                Service = "OneDrive for Business"
                ActiveUsers = ($oneDriveReport | Where-Object { $_.ViewedOrEditedFileCount -gt 0 }).Count
                TotalUsers = $oneDriveReport.Count
                UsagePercentage = [math]::Round((($oneDriveReport | Where-Object { $_.ViewedOrEditedFileCount -gt 0 }).Count / $oneDriveReport.Count) * 100, 2)
            }
        }
        
        # Get Exchange user activity
        if ($IncludeExchange) {
            Write-Host "Retrieving Exchange Online usage data..." -ForegroundColor Cyan
            $exchangeReport = Get-MailboxUsageDetailReport -Period $Days
            $reports += [PSCustomObject]@{
                Service = "Exchange Online"
                ActiveUsers = ($exchangeReport | Where-Object { $_.SendCount -gt 0 -or $_.ReceiveCount -gt 0 }).Count
                TotalUsers = $exchangeReport.Count
                UsagePercentage = [math]::Round((($exchangeReport | Where-Object { $_.SendCount -gt 0 -or $_.ReceiveCount -gt 0 }).Count / $exchangeReport.Count) * 100, 2)
            }
        }
        
        # Get SharePoint user activity
        if ($IncludeSharePoint) {
            Write-Host "Retrieving SharePoint Online usage data..." -ForegroundColor Cyan
            $sharePointReport = Get-SPOUserActivityUserDetail -Period $Days
            $reports += [PSCustomObject]@{
                Service = "SharePoint Online"
                ActiveUsers = ($sharePointReport | Where-Object { $_.ViewedOrEditedFileCount -gt 0 }).Count
                TotalUsers = $sharePointReport.Count
                UsagePercentage = [math]::Round((($sharePointReport | Where-Object { $_.ViewedOrEditedFileCount -gt 0 }).Count / $sharePointReport.Count) * 100, 2)
            }
        }
        
        # Export to CSV
        $reports | Export-Csv -Path $OutputPath -NoTypeInformation
        Write-Host "Successfully generated M365 usage report at $OutputPath" -ForegroundColor Green
        
        # Display summary
        Write-Host "`nM365 Usage Report Summary (Last $Days days):" -ForegroundColor Cyan
        $reports | Format-Table -AutoSize
        
        return $reports
    }
    catch {
        Write-Error "Error generating usage report: $_"
    }
}

# Function to create a conditional access policy
function New-M365ConditionalAccessPolicy {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$PolicyName,
        
        [Parameter()]
        [string[]]$IncludeUsers,
        
        [Parameter()]
        [string[]]$ExcludeUsers,
        
        [Parameter()]
        [string[]]$IncludeGroups,
        
        [Parameter()]
        [string[]]$IncludeApplications,
        
        [Parameter()]
        [ValidateSet("all", "browser", "mobileAppsAndDesktopClients", "exchangeActiveSync", "other")]
        [string[]]$ClientAppTypes = @("all"),
        
        [Parameter()]
        [ValidateSet("require", "block")]
        [string]$AccessControl = "require",
        
        [Parameter()]
        [switch]$RequireMFA,
        
        [Parameter()]
        [switch]$BlockLegacyAuth,
        
        [Parameter()]
        [switch]$RequireCompliantDevice,
        
        [Parameter()]
        [switch]$RequireApprovedApp
    )
    
    try {
        # Check if AzureAD module is available
        if (-not (Get-Module -ListAvailable -Name AzureADPreview)) {
            Write-Error "AzureADPreview module is required for Conditional Access policies. Please install it using: Install-Module AzureADPreview -Force"
            return
        }
        
        # Check if connected to Azure AD Preview
        try {
            Get-AzureADMSConditionalAccessPolicy -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Error "Not connected to AzureAD Preview. Please run Connect-AzureAD first with the Preview module."
            return
        }
        
        # Create conditions
        $conditions = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessConditionSet
        
        # Configure users
        $conditions.Users = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessUserCondition
        $conditions.Users.IncludeUsers = if ($IncludeUsers) { $IncludeUsers } else { @("All") }
        if ($ExcludeUsers) {
            $conditions.Users.ExcludeUsers = $ExcludeUsers
        }
        if ($IncludeGroups) {
            $conditions.Users.IncludeGroups = $IncludeGroups
        }
        
        # Configure applications
        $conditions.Applications = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessApplicationCondition
        $conditions.Applications.IncludeApplications = if ($IncludeApplications) { $IncludeApplications } else { @("All") }
        
        # Configure client app types
        $conditions.ClientAppTypes = $ClientAppTypes
        
        # Create grant controls
        $grantControls = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessGrantControls
        $grantControls.Operator = "OR"
        $grantControls.BuiltInControls = @()
        
        if ($RequireMFA) {
            $grantControls.BuiltInControls += "mfa"
        }
        
        if ($RequireCompliantDevice) {
            $grantControls.BuiltInControls += "compliantDevice"
        }
        
        if ($RequireApprovedApp) {
            $grantControls.BuiltInControls += "approvedApplication"
        }
        
        # Create session controls for legacy auth blocking
        $sessionControls = $null
        if ($BlockLegacyAuth -and $ClientAppTypes -contains "exchangeActiveSync") {
            $sessionControls = New-Object -TypeName Microsoft.Open.MSGraph.Model.ConditionalAccessSessionControls
            $sessionControls.ApplicationEnforcedRestrictions = New-Object -TypeName Microsoft.Open.MSGraph.Model.ApplicationEnforcedRestrictionsSessionControl
            $sessionControls.ApplicationEnforcedRestrictions.IsEnabled = $true
        }
        
        # Set state based on access control
        $state = if ($AccessControl -eq "require") { "enabled" } else { "disabled" }
        
        # Create policy
        $params = @{
            DisplayName = $PolicyName
            State = $state
            Conditions = $conditions
            GrantControls = $grantControls
        }
        
        if ($sessionControls) {
            $params.SessionControls = $sessionControls
        }
        
        $policy = New-AzureADMSConditionalAccessPolicy @params
        
        Write-Host "Successfully created Conditional Access policy: $PolicyName" -ForegroundColor Green
        return $policy
    }
    catch {
        Write-Error "Error creating Conditional Access policy: $_"
    }
}

# Sample usage examples
function Show-M365ToolkitExamples {
    Write-Host "Example 1: Connect to Microsoft 365 services" -ForegroundColor Cyan
    Write-Host 'Connect-M365Services -Credential (Get-Credential) -ExchangeOnline -AzureAD -SharePointOnline -MicrosoftTeams' -ForegroundColor Yellow
    
    Write-Host "`nExample 2: Create users in bulk from CSV file" -ForegroundColor Cyan
    Write-Host 'New-M365BulkUsers -CsvPath "C:\Users.csv" -DefaultPassword "Password123!" -LicenseSkuId "tenant:ENTERPRISEPACK"' -ForegroundColor Yellow
    
    Write-Host "`nExample 3: Export user license information" -ForegroundColor Cyan
    Write-Host 'Export-M365UserLicenses -OutputPath "C:\M365UserLicenses.csv" -IncludeDisabledUsers' -ForegroundColor Yellow
    
    Write-Host "`nExample 4: Create a new Teams team with channels" -ForegroundColor Cyan
    Write-Host 'New-M365TeamWithChannels -TeamName "Marketing Team" -TeamDescription "Team for marketing department" -Channels @("General", "Campaigns", "Social Media") -OwnerEmails @("admin@contoso.com") -MemberEmails @("user1@contoso.com", "user2@contoso.com")' -ForegroundColor Yellow
    
    Write-Host "`nExample 5: Manage SharePoint site permissions" -ForegroundColor Cyan
    Write-Host 'Set-SPOSitePermissions -SiteUrl "https://contoso.sharepoint.com/sites/Marketing" -OwnersToAdd @{"John Doe" = "john@contoso.com"} -MembersToAdd @{"Jane Smith" = "jane@contoso.com"} -VisitorsToRemove @("formeruser@contoso.com")' -ForegroundColor Yellow
    
    Write-Host "`nExample 6: Generate Microsoft 365 usage report" -ForegroundColor Cyan
    Write-Host 'Get-M365UsageReport -OutputPath "C:\M365UsageReport.csv" -Days 30 -IncludeTeams -IncludeOneDrive -IncludeExchange -IncludeSharePoint' -ForegroundColor Yellow
    
    Write-Host "`nExample 7: Create a conditional access policy requiring MFA" -ForegroundColor Cyan
    Write-Host 'New-M365ConditionalAccessPolicy -PolicyName "Require MFA for All Users" -RequireMFA' -ForegroundColor Yellow
}
