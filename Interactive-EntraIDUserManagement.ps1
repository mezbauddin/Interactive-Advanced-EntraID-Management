<#
.SCRIPT NAME
    Interactive Advanced Entra ID User Management Script

.SYNOPSIS
    PowerShell script to interactively manage users, groups, licenses, and roles in Entra ID (Azure Active Directory).

.DESCRIPTION
    This interactive script provides functionality to:
    1. List all users with detailed properties.
    2. Add a new user and assign licenses and roles.
    3. Update user properties and manage group memberships interactively.
    4. Assign or remove licenses interactively.
    5. Manage MFA settings for users.

.AUTHOR
    Mezba Uddin

.VERSION
    2.2

.LASTUPDATED
    2024-11-06

.NOTES
    - Requires the Microsoft.Graph module.
    - Run `Connect-MgGraph` to authenticate before executing the script.
    - Ensure the account running the script has appropriate permissions.

#>

# Import the Microsoft.Graph module (install if not already installed)
Write-Host "Checking for Microsoft.Graph module..." -ForegroundColor Cyan
$moduleInstalled = Get-Module -Name Microsoft.Graph -ListAvailable
if (-not $moduleInstalled) {
    Write-Host "Microsoft.Graph module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name Microsoft.Graph -Force -Scope CurrentUser -ErrorAction Stop
        Write-Host "Microsoft.Graph module installed successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Error installing Microsoft.Graph module: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please try installing the module manually using: Install-Module -Name Microsoft.Graph -Force -Scope CurrentUser" -ForegroundColor Yellow
        exit
    }
}
else {
    Write-Host "Microsoft.Graph module is already installed." -ForegroundColor Green
}

Write-Host "Importing Microsoft.Graph module..." -ForegroundColor Cyan
try {
    Import-Module Microsoft.Graph -ErrorAction Stop
    Write-Host "Microsoft.Graph module imported successfully!" -ForegroundColor Green
}
catch {
    Write-Host "Error importing Microsoft.Graph module: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Please try importing the module manually using: Import-Module Microsoft.Graph" -ForegroundColor Yellow
    exit
}

# Function to ensure Graph connection
function Connect-EntraGraph {
    try {
        $context = Get-MgContext
        if (-not $context) {
            Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan
            Connect-MgGraph -Scopes @(
                "User.ReadWrite.All",
                "Group.ReadWrite.All",
                "Directory.AccessAsUser.All",
                "UserAuthenticationMethod.ReadWrite.All"
            ) -ErrorAction Stop
            Write-Host "Successfully connected to Microsoft Graph" -ForegroundColor Green
            return $true
        }
        
        # Check if we have all required scopes
        $requiredScopes = @(
            "User.ReadWrite.All",
            "Group.ReadWrite.All",
            "Directory.AccessAsUser.All",
            "UserAuthenticationMethod.ReadWrite.All"
        )
        
        $missingScopes = $requiredScopes | Where-Object { $context.Scopes -notcontains $_ }
        if ($missingScopes) {
            Write-Host "`nReconnecting to Microsoft Graph to acquire additional permissions..." -ForegroundColor Cyan
            Disconnect-MgGraph
            Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
            Write-Host "Successfully reconnected to Microsoft Graph with updated permissions" -ForegroundColor Green
        }
        return $true
    }
    catch {
        Write-Host "Error connecting to Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Failed to connect to Microsoft Graph" -ErrorDetails $_.Exception.Message
        return $false
    }
}

# Function to display the banner
function Show-Banner {
    $banner = @"
    ╔══════════════════════════════════════════════════════════════════╗
    ║                                                                  ║
    ║                         MEZBA UDDIN                              ║
    ║                                                                  ║
    ║              Microsoft Most Valuable Professional                ║
    ║                           (MVP)                                  ║
    ║                                                                  ║
    ╠══════════════════════════════════════════════════════════════════╣
    ║                                                                  ║
    ║ 💡 Inspiring IT Innovation and Transformation                    ║
    ║ 🌐 Website:    mezbauddin.com                                   ║
    ║ 🔗 LinkedIn:   linkedin.com/in/mezbauddin                       ║
    ║ 📧 Email:      contact@mezbauddin.com                           ║
    ║                                                                  ║
    ╚══════════════════════════════════════════════════════════════════╝

"@
    $title = @"
    ╔══════════════════════════════════════════════════════════════════╗
    ║                                                                  ║
    ║                ENTRA ID USER MANAGEMENT UTILITY                  ║
    ║                                                                  ║
    ╚══════════════════════════════════════════════════════════════════╝

"@
    Write-Host $banner -ForegroundColor Magenta
    Write-Host $title -ForegroundColor Cyan
    Write-Host "    Version: 2.2 | Last Updated: 2024-11-06" -ForegroundColor Gray
    Write-Host "    Powered by Microsoft Graph API`n" -ForegroundColor Gray
}

# Function to display main menu
function Show-Menu {
    Write-Host "`n    ╔══════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "    ║           AVAILABLE OPTIONS          ║" -ForegroundColor Green
    Write-Host "    ╠══════════════════════════════════════╣" -ForegroundColor Green
    Write-Host "    ║                                      ║" -ForegroundColor Green
    Write-Host "    ║    1. List all users                ║" -ForegroundColor Cyan
    Write-Host "    ║    2. Add a new user                ║" -ForegroundColor Cyan
    Write-Host "    ║    3. Update a user                 ║" -ForegroundColor Cyan
    Write-Host "    ║    4. License Management            ║" -ForegroundColor Cyan
    Write-Host "    ║    5. MFA Management                ║" -ForegroundColor Cyan
    Write-Host "    ║    6. Exit                          ║" -ForegroundColor Red
    Write-Host "    ║                                      ║" -ForegroundColor Green
    Write-Host "    ╚══════════════════════════════════════╝" -ForegroundColor Green
    Write-Host "`n    Select an option (1-6): " -ForegroundColor Yellow -NoNewline
}

# Function to show license management menu
function Show-LicenseMenu {
    Write-Host "`n    ╔══════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "    ║        LICENSE MANAGEMENT            ║" -ForegroundColor Green
    Write-Host "    ╠══════════════════════════════════════╣" -ForegroundColor Green
    Write-Host "    ║                                      ║" -ForegroundColor Green
    Write-Host "    ║    1. Manage Single User            ║" -ForegroundColor Cyan
    Write-Host "    ║    2. Manage Multiple Users         ║" -ForegroundColor Cyan
    Write-Host "    ║    3. Back to Main Menu             ║" -ForegroundColor Yellow
    Write-Host "    ║                                      ║" -ForegroundColor Green
    Write-Host "    ╚══════════════════════════════════════╝" -ForegroundColor Green
    Write-Host "`n    Enter your choice (1-3): " -ForegroundColor Yellow -NoNewline

    $choice = Read-Host
    
    switch ($choice) {
        "1" { Manage-Licenses }
        "2" { Manage-BulkLicenses }
        "3" { return }
        default { 
            Write-Host "`n    Invalid choice. Please try again." -ForegroundColor Yellow
            Show-LicenseMenu
        }
    }
}

# Function to list all users
function List-AllUsers {
    Write-Host "Fetching all Entra ID users..."
    $users = Get-MgUser -All
    foreach ($user in $users) {
        Write-Host "Display Name: $($user.DisplayName), Email: $($user.Mail), ID: $($user.Id)"
    }
}

# Function to write error logs
function Write-ErrorLog {
    param(
        [string]$ErrorMessage,
        [string]$ErrorDetails
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] ERROR: $ErrorMessage`nDetails: $ErrorDetails`n"
    
    try {
        $logPath = Join-Path $PSScriptRoot "license_management.log"
        Add-Content -Path $logPath -Value $logMessage
    }
    catch {
        Write-Host "Failed to write to log file: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to validate user input
function Test-UserInput {
    param(
        [string]$UserPrincipalName
    )
    
    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        throw "UserPrincipalName cannot be empty."
    }
    
    if (-not ($UserPrincipalName -match "^[^@]+@[^@]+\.[^@]+$")) {
        throw "Invalid email format. Please enter a valid email address."
    }
}

# Function to get and display available licenses
function Get-AvailableLicenses {
    try {
        $subscriptions = Get-MgSubscribedSku -ErrorAction Stop
        if (-not $subscriptions) {
            Write-Host "No licenses found in the tenant." -ForegroundColor Yellow
            return $null
        }

        $i = 1
        $licenseOptions = @{}
        
        Write-Host "`nAvailable Licenses:"
        Write-Host "==================="
        foreach ($sub in $subscriptions) {
            $availableUnits = $sub.PrepaidUnits.Enabled - $sub.ConsumedUnits
            Write-Host "$i. $($sub.SkuPartNumber) - Available: $availableUnits" -ForegroundColor Cyan
            $licenseOptions[$i] = @{
                SkuId = $sub.SkuId
                AvailableUnits = $availableUnits
                SkuPartNumber = $sub.SkuPartNumber
            }
            $i++
        }
        Write-Host "==================="
        
        do {
            $selection = Read-Host "Select a license number (or press Enter to skip)"
            if ([string]::IsNullOrEmpty($selection)) { return $null }
            
            if ($licenseOptions[$selection]) {
                if ($licenseOptions[$selection].AvailableUnits -le 0) {
                    Write-Host "No available licenses for $($licenseOptions[$selection].SkuPartNumber)" -ForegroundColor Yellow
                    return $null
                }
                return $licenseOptions[$selection].SkuId
            }
            Write-Host "Invalid selection. Please try again." -ForegroundColor Yellow
        } while ($true)
    }
    catch {
        Write-Host "Error retrieving licenses: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Failed to retrieve licenses" -ErrorDetails $_.Exception.Message
        return $null
    }
}

# Function to view user's current licenses
function Show-UserLicenses {
    param (
        [Parameter(Mandatory=$true)]
        [string]$UserId
    )
    
    try {
        $userLicenses = Get-MgUserLicenseDetail -UserId $UserId -ErrorAction Stop
        if ($userLicenses.Count -eq 0) {
            Write-Host "User has no licenses assigned." -ForegroundColor Yellow
            return $false
        }
        
        Write-Host "`nUser's current licenses:"
        Write-Host "======================="
        foreach ($license in $userLicenses) {
            Write-Host "`nLicense: $($license.SkuPartNumber)" -ForegroundColor Cyan
            Write-Host "Status: Assigned" -ForegroundColor Gray
            
            # Display service plans
            Write-Host "Enabled Service Plans:" -ForegroundColor Gray
            $license.ServicePlans | Where-Object { $_.ProvisioningStatus -eq "Success" } | ForEach-Object {
                Write-Host "  - $($_.ServicePlanName)" -ForegroundColor Green
            }
            
            Write-Host "Disabled Service Plans:" -ForegroundColor Gray
            $license.ServicePlans | Where-Object { $_.ProvisioningStatus -ne "Success" } | ForEach-Object {
                Write-Host "  - $($_.ServicePlanName)" -ForegroundColor Yellow
            }
        }
        Write-Host "======================="
        return $true
    }
    catch {
        Write-Host "Error retrieving user licenses: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Failed to retrieve user licenses" -ErrorDetails $_.Exception.Message
        return $false
    }
}

# Function to add a new user
function Add-NewUser {
    try {
        $DisplayName = Read-Host "Enter the display name for the new user"
        $UserPrincipalName = Read-Host "Enter the user principal name (e.g., john.doe@yourdomain.com)"
        
        # Validate UPN format
        if (-not ($UserPrincipalName -match "^[^@]+@[^@]+\.[^@]+$")) {
            Write-Host "Invalid email format. Please enter a valid email address." -ForegroundColor Red
            return
        }
        
        $MailNickname = Read-Host "Enter the mail nickname for the user"
        $Password = Read-Host "Enter a temporary password for the user (password will need to be reset on first login)"
        
        # Validate password complexity
        if ($Password.Length -lt 8) {
            Write-Host "Password must be at least 8 characters long." -ForegroundColor Red
            return
        }

        Write-Host "`nSelect a license to assign:"
        $LicenseSkuId = Get-AvailableLicenses
        $RoleId = Read-Host "Enter the Directory Role ID (Leave blank to skip role assignment)"

        Write-Host "Adding new user: $DisplayName..." -ForegroundColor Cyan
        $newUser = New-MgUser -AccountEnabled $true `
            -DisplayName $DisplayName `
            -UserPrincipalName $UserPrincipalName `
            -MailNickname $MailNickname `
            -PasswordProfile @{
                Password = $Password
                ForceChangePasswordNextSignIn = $true
                ForceChangePasswordNextSignInWithMfa = $false
            } -ErrorAction Stop

        Write-Host "User $DisplayName added successfully!" -ForegroundColor Green

        if ($LicenseSkuId) {
            try {
                Add-MgUserLicense -UserId $newUser.Id -AddLicenses @{SkuId=$LicenseSkuId} -RemoveLicenses @() -ErrorAction Stop
                Write-Host "License assigned successfully" -ForegroundColor Green
            }
            catch {
                Write-Host "Error assigning license: $($_.Exception.Message)" -ForegroundColor Red
                Write-ErrorLog -ErrorMessage "Failed to assign license to new user" -ErrorDetails $_.Exception.Message
            }
        }

        if ($RoleId) {
            try {
                New-MgDirectoryRoleMember -DirectoryRoleId $RoleId -DirectoryObjectId $newUser.Id -ErrorAction Stop
                Write-Host "Role assigned: $RoleId" -ForegroundColor Green
            }
            catch {
                Write-Host "Error assigning role: $($_.Exception.Message)" -ForegroundColor Red
                Write-ErrorLog -ErrorMessage "Failed to assign role to new user" -ErrorDetails $_.Exception.Message
            }
        }
    }
    catch {
        Write-Host "Error creating user: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Failed to create new user" -ErrorDetails $_.Exception.Message
    }
}

# Function to update an existing user
function Update-User {
    try {
        Write-Host "`nUser Selection" -ForegroundColor Green
        Write-Host "=============" -ForegroundColor Green
        $user = Find-EntraIDUser
        
        if (-not $user) {
            Write-Host "No user selected. Returning to main menu..." -ForegroundColor Yellow
            return
        }

        Write-Host "`nCurrent User Details:" -ForegroundColor Cyan
        Write-Host "Display Name: $($user.DisplayName)" -ForegroundColor Gray
        Write-Host "UPN: $($user.UserPrincipalName)" -ForegroundColor Gray
        Write-Host "Job Title: $($user.JobTitle)" -ForegroundColor Gray
        Write-Host "Department: $($user.Department)" -ForegroundColor Gray

        $NewDisplayName = Read-Host "Enter the new display name (Leave blank to keep unchanged)"
        $NewJobTitle = Read-Host "Enter the new job title (Leave blank to keep unchanged)"
        $NewDepartment = Read-Host "Enter the new department (Leave blank to keep unchanged)"
        $AddGroups = Read-Host "Enter Group IDs to add the user to (comma-separated, or leave blank)"
        $RemoveGroups = Read-Host "Enter Group IDs to remove the user from (comma-separated, or leave blank)"

        $updateParams = @{}
        if ($NewDisplayName) { $updateParams.DisplayName = $NewDisplayName }
        if ($NewJobTitle) { $updateParams.JobTitle = $NewJobTitle }
        if ($NewDepartment) { $updateParams.Department = $NewDepartment }

        if ($updateParams.Count -gt 0) {
            try {
                Update-MgUser -UserId $user.Id @updateParams -ErrorAction Stop
                Write-Host "User properties updated successfully!" -ForegroundColor Green
            }
            catch {
                Write-Host "Error updating user properties: $($_.Exception.Message)" -ForegroundColor Red
                Write-ErrorLog -ErrorMessage "Failed to update user properties" -ErrorDetails $_.Exception.Message
            }
        }

        if ($AddGroups) {
            $groupIdsToAdd = $AddGroups -split "," | ForEach-Object { $_.Trim() }
            foreach ($groupId in $groupIdsToAdd) {
                try {
                    # Verify group exists
                    $group = Get-MgGroup -GroupId $groupId -ErrorAction Stop
                    Add-MgGroupMember -GroupId $groupId -DirectoryObjectId $user.Id -ErrorAction Stop
                    Write-Host "Added to group: $($group.DisplayName)" -ForegroundColor Green
                }
                catch {
                    Write-Host "Error adding to group $groupId : $($_.Exception.Message)" -ForegroundColor Red
                    Write-ErrorLog -ErrorMessage "Failed to add user to group" -ErrorDetails $_.Exception.Message
                }
            }
        }

        if ($RemoveGroups) {
            $groupIdsToRemove = $RemoveGroups -split "," | ForEach-Object { $_.Trim() }
            foreach ($groupId in $groupIdsToRemove) {
                try {
                    # Verify group exists
                    $group = Get-MgGroup -GroupId $groupId -ErrorAction Stop
                    Remove-MgGroupMember -GroupId $groupId -DirectoryObjectId $user.Id -ErrorAction Stop
                    Write-Host "Removed from group: $($group.DisplayName)" -ForegroundColor Green
                }
                catch {
                    Write-Host "Error removing from group $groupId : $($_.Exception.Message)" -ForegroundColor Red
                    Write-ErrorLog -ErrorMessage "Failed to remove user from group" -ErrorDetails $_.Exception.Message
                }
            }
        }
    }
    catch {
        Write-Host "An unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Unexpected error in Update-User" -ErrorDetails $_.Exception.Message
    }
}

# Function to search and select a user
function Find-EntraIDUser {
    try {
        # Ensure we're connected before proceeding
        if (-not (Connect-EntraGraph)) {
            Write-Host "Unable to proceed without Microsoft Graph connection." -ForegroundColor Red
            return $null
        }

        do {
            Write-Host "`nSearch Options:" -ForegroundColor Cyan
            Write-Host "1. Search by name or email"
            Write-Host "2. List all users"
            Write-Host "3. Back to previous menu"
            $searchChoice = Read-Host "`nEnter your choice (1-3)"

            switch ($searchChoice) {
                "1" {
                    $searchTerm = Read-Host "Enter part of name or email to search"
                    if ([string]::IsNullOrWhiteSpace($searchTerm)) {
                        Write-Host "Search term cannot be empty" -ForegroundColor Yellow
                        continue
                    }

                    Write-Host "`nSearching for users matching '$searchTerm'..." -ForegroundColor Cyan
                    try {
                        # Using startsWith for more reliable search
                        $users = Get-MgUser -Filter "startsWith(displayName,'$searchTerm') or startsWith(userPrincipalName,'$searchTerm')" -Top 10 -ErrorAction Stop
                        
                        if (-not $users) {
                            # If no results with startsWith, try searching with endsWith
                            $users = Get-MgUser -Filter "endsWith(displayName,'$searchTerm') or endsWith(userPrincipalName,'$searchTerm')" -Top 10 -ErrorAction Stop
                        }
                        
                        if (-not $users) {
                            Write-Host "No users found matching '$searchTerm'" -ForegroundColor Yellow
                            
                            # Offer to list all users if no matches found
                            $listAll = Read-Host "Would you like to see all users instead? (Y/N)"
                            if ($listAll -eq 'Y' -or $listAll -eq 'y') {
                                $users = Get-MgUser -Top 20 -ErrorAction Stop
                            } else {
                                continue
                            }
                        }
                    }
                    catch {
                        Write-Host "Error in search query: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host "Retrieving all users instead..." -ForegroundColor Yellow
                        $users = Get-MgUser -Top 20 -ErrorAction Stop
                    }

                    Write-Host "`nFound Users:" -ForegroundColor Green
                    Write-Host "=============" -ForegroundColor Green
                    $userList = @{}
                    $i = 1
                    foreach ($user in $users) {
                        Write-Host "`n$i. Display Name: $($user.DisplayName)" -ForegroundColor Cyan
                        Write-Host "   Email: $($user.UserPrincipalName)" -ForegroundColor Gray
                        if ($user.Department) {
                            Write-Host "   Department: $($user.Department)" -ForegroundColor Gray
                        }
                        if ($user.JobTitle) {
                            Write-Host "   Job Title: $($user.JobTitle)" -ForegroundColor Gray
                        }
                        $userList[$i] = $user
                        $i++
                    }
                    Write-Host "`n=============" -ForegroundColor Green

                    $selection = Read-Host "`nEnter the number of the user to select (or press Enter to search again)"
                    if (-not [string]::IsNullOrWhiteSpace($selection) -and $userList.ContainsKey([int]$selection)) {
                        return $userList[[int]$selection]
                    }
                }
                "2" {
                    Write-Host "`nRetrieving all users (limited to 20)..." -ForegroundColor Cyan
                    $users = Get-MgUser -Top 20 -ErrorAction Stop
                    
                    Write-Host "`nAll Users:" -ForegroundColor Green
                    Write-Host "==========" -ForegroundColor Green
                    $userList = @{}
                    $i = 1
                    foreach ($user in $users) {
                        Write-Host "`n$i. Display Name: $($user.DisplayName)" -ForegroundColor Cyan
                        Write-Host "   Email: $($user.UserPrincipalName)" -ForegroundColor Gray
                        if ($user.Department) {
                            Write-Host "   Department: $($user.Department)" -ForegroundColor Gray
                        }
                        if ($user.JobTitle) {
                            Write-Host "   Job Title: $($user.JobTitle)" -ForegroundColor Gray
                        }
                        $userList[$i] = $user
                        $i++
                    }
                    Write-Host "`n==========" -ForegroundColor Green

                    $selection = Read-Host "`nEnter the number of the user to select (or press Enter to go back)"
                    if (-not [string]::IsNullOrWhiteSpace($selection) -and $userList.ContainsKey([int]$selection)) {
                        return $userList[[int]$selection]
                    }
                }
                "3" {
                    return $null
                }
                default {
                    Write-Host "Invalid choice. Please enter a number between 1 and 3." -ForegroundColor Yellow
                }
            }
        } while ($true)
    }
    catch {
        Write-Host "Error searching for users: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Failed to search users" -ErrorDetails $_.Exception.Message
        return $null
    }
}

# Function to assign or remove licenses
function Manage-Licenses {
    try {
        # Ensure we're connected before proceeding
        if (-not (Connect-EntraGraph)) {
            Write-Host "Unable to proceed without Microsoft Graph connection." -ForegroundColor Red
            return
        }

        Write-Host "`nUser Selection" -ForegroundColor Green
        Write-Host "=============" -ForegroundColor Green
        $user = Find-EntraIDUser
        
        if (-not $user) {
            Write-Host "No user selected. Returning to main menu..." -ForegroundColor Yellow
            return
        }

        do {
            Write-Host "`nLicense Management Menu for $($user.DisplayName)" -ForegroundColor Green
            Write-Host "=====================" -ForegroundColor Green
            Write-Host "1. Add License" -ForegroundColor Cyan
            Write-Host "2. Remove License" -ForegroundColor Cyan
            Write-Host "3. View Current Licenses" -ForegroundColor Cyan
            Write-Host "4. Select Different User" -ForegroundColor Cyan
            Write-Host "5. Back to Main Menu" -ForegroundColor Cyan
            Write-Host "=====================" -ForegroundColor Green
            
            $choice = Read-Host "Enter your choice (1-5)"

            switch ($choice) {
                "1" {
                    Write-Host "`nSelect a license to assign:"
                    $LicenseSkuId = Get-AvailableLicenses
                    if ($LicenseSkuId) {
                        try {
                            Add-MgUserLicense -UserId $user.Id -AddLicenses @{SkuId=$LicenseSkuId} -RemoveLicenses @() -ErrorAction Stop
                            Write-Host "License assigned successfully to $($user.UserPrincipalName)" -ForegroundColor Green
                        }
                        catch {
                            Write-Host "Error assigning license: $($_.Exception.Message)" -ForegroundColor Red
                            Write-ErrorLog -ErrorMessage "Failed to assign license" -ErrorDetails $_.Exception.Message
                        }
                    }
                }
                "2" {
                    try {
                        $userLicenses = Get-MgUserLicenseDetail -UserId $user.Id -ErrorAction Stop
                        if ($userLicenses.Count -eq 0) {
                            Write-Host "User has no licenses assigned." -ForegroundColor Yellow
                            break
                        }
                        
                        Write-Host "`nUser's current licenses:"
                        Write-Host "======================="
                        for ($i = 0; $i -lt $userLicenses.Count; $i++) {
                            Write-Host "$($i+1). $($userLicenses[$i].SkuPartNumber)" -ForegroundColor Cyan
                        }
                        Write-Host "======================="
                        
                        $selection = Read-Host "Select the license number to remove"
                        if ($selection -ge 1 -and $selection -le $userLicenses.Count) {
                            try {
                                $LicenseSkuId = $userLicenses[$selection-1].SkuId
                                Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($LicenseSkuId) -ErrorAction Stop
                                Write-Host "License removed successfully from $($user.UserPrincipalName)" -ForegroundColor Green
                                
                                # Verify the license removal
                                $updatedLicenses = Get-MgUserLicenseDetail -UserId $user.Id
                                if ($updatedLicenses.SkuId -notcontains $LicenseSkuId) {
                                    Write-Host "Verified: License removal confirmed" -ForegroundColor Green
                                } else {
                                    Write-Host "Warning: License might not have been removed properly. Please verify." -ForegroundColor Yellow
                                }
                            }
                            catch {
                                Write-Host "Error removing license: $($_.Exception.Message)" -ForegroundColor Red
                                Write-ErrorLog -ErrorMessage "Failed to remove license" -ErrorDetails $_.Exception.Message
                            }
                        }
                        else {
                            Write-Host "Invalid selection." -ForegroundColor Yellow
                        }
                    }
                    catch {
                        Write-Host "Error retrieving user licenses: $($_.Exception.Message)" -ForegroundColor Red
                        Write-ErrorLog -ErrorMessage "Failed to retrieve user licenses" -ErrorDetails $_.Exception.Message
                    }
                }
                "3" {
                    Show-UserLicenses -UserId $user.Id
                }
                "4" {
                    $user = Find-EntraIDUser
                    if (-not $user) {
                        Write-Host "No user selected. Returning to main menu..." -ForegroundColor Yellow
                        return
                    }
                }
                "5" {
                    Write-Host "Returning to main menu..." -ForegroundColor Green
                    return
                }
                default {
                    Write-Host "Invalid choice. Please enter a number between 1 and 5." -ForegroundColor Yellow
                }
            }
        } while ($true)
    }
    catch [System.Management.Automation.ParameterBindingException] {
        Write-Host "Invalid input: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Invalid input parameter" -ErrorDetails $_.Exception.Message
    }
    catch {
        Write-Host "An unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Unexpected error in license management" -ErrorDetails $_.Exception.Message
    }
}

# Function to perform bulk license operations
function Manage-BulkLicenses {
    try {
        # Ensure we're connected before proceeding
        if (-not (Connect-EntraGraph)) {
            Write-Host "Unable to proceed without Microsoft Graph connection." -ForegroundColor Red
            return
        }

        Write-Host "`nBulk License Operations" -ForegroundColor Green
        Write-Host "=====================" -ForegroundColor Green
        Write-Host "1. Add license to multiple users" -ForegroundColor Cyan
        Write-Host "2. Remove license from multiple users" -ForegroundColor Cyan
        Write-Host "3. Back to main menu" -ForegroundColor Cyan
        Write-Host "=====================" -ForegroundColor Green
        
        $choice = Read-Host "Enter your choice (1-3)"

        switch ($choice) {
            "1" {
                Write-Host "`nSelect the license to assign:" -ForegroundColor Cyan
                $LicenseSkuId = Get-AvailableLicenses
                if (-not $LicenseSkuId) { return }

                $selectedUsers = @()
                do {
                    $user = Find-EntraIDUser
                    if ($user) {
                        $selectedUsers += $user
                        Write-Host "Added $($user.DisplayName) to selection" -ForegroundColor Green
                    }
                    $continue = Read-Host "Add another user? (Y/N)"
                } while ($continue -eq 'Y' -or $continue -eq 'y')

                if ($selectedUsers.Count -eq 0) {
                    Write-Host "No users selected." -ForegroundColor Yellow
                    return
                }

                Write-Host "`nAssigning license to selected users..." -ForegroundColor Cyan
                foreach ($user in $selectedUsers) {
                    try {
                        Add-MgUserLicense -UserId $user.Id -AddLicenses @{SkuId=$LicenseSkuId} -RemoveLicenses @() -ErrorAction Stop
                        Write-Host "License assigned successfully to $($user.UserPrincipalName)" -ForegroundColor Green
                    }
                    catch {
                        Write-Host "Error assigning license to $($user.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Red
                        Write-ErrorLog -ErrorMessage "Failed to assign license to user" -ErrorDetails $_.Exception.Message
                    }
                }
            }
            "2" {
                Write-Host "`nSelect users to remove license from:" -ForegroundColor Cyan
                $selectedUsers = @()
                do {
                    $user = Find-EntraIDUser
                    if ($user) {
                        $userLicenses = Get-MgUserLicenseDetail -UserId $user.Id
                        if ($userLicenses.Count -eq 0) {
                            Write-Host "$($user.DisplayName) has no licenses." -ForegroundColor Yellow
                            continue
                        }
                        Write-Host "`nLicenses for $($user.DisplayName):" -ForegroundColor Cyan
                        for ($i = 0; $i -lt $userLicenses.Count; $i++) {
                            Write-Host "$($i+1). $($userLicenses[$i].SkuPartNumber)" -ForegroundColor Gray
                        }
                        $licenseChoice = Read-Host "Select license number to remove"
                        if ($licenseChoice -ge 1 -and $licenseChoice -le $userLicenses.Count) {
                            $selectedUsers += @{
                                User = $user
                                LicenseSkuId = $userLicenses[$licenseChoice-1].SkuId
                            }
                            Write-Host "Added $($user.DisplayName) to selection" -ForegroundColor Green
                        }
                    }
                    $continue = Read-Host "Add another user? (Y/N)"
                } while ($continue -eq 'Y' -or $continue -eq 'y')

                if ($selectedUsers.Count -eq 0) {
                    Write-Host "No users selected." -ForegroundColor Yellow
                    return
                }

                Write-Host "`nRemoving licenses from selected users..." -ForegroundColor Cyan
                foreach ($selected in $selectedUsers) {
                    try {
                        Set-MgUserLicense -UserId $selected.User.Id -AddLicenses @() -RemoveLicenses @($selected.LicenseSkuId) -ErrorAction Stop
                        Write-Host "License removed successfully from $($selected.User.UserPrincipalName)" -ForegroundColor Green
                    }
                    catch {
                        Write-Host "Error removing license from $($selected.User.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Red
                        Write-ErrorLog -ErrorMessage "Failed to remove license from user" -ErrorDetails $_.Exception.Message
                    }
                }
            }
            "3" {
                return
            }
            default {
                Write-Host "Invalid choice. Please enter a number between 1 and 3." -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Host "An error occurred during bulk license operation: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Error in bulk license operation" -ErrorDetails $_.Exception.Message
    }
}

# Function to manage MFA
function Manage-MFA {
    try {
        # Ensure we're connected before proceeding
        if (-not (Connect-EntraGraph)) {
            Write-Host "Unable to proceed without Microsoft Graph connection." -ForegroundColor Red
            return
        }

        Write-Host "`nMFA Management" -ForegroundColor Green
        Write-Host "=============" -ForegroundColor Green
        
        # Find user
        $user = Find-EntraIDUser
        if (-not $user) {
            Write-Host "No user selected. Returning to main menu..." -ForegroundColor Yellow
            return
        }

        try {
            do {
                $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction Stop
                
                Write-Host "`nCurrent Authentication Methods for $($user.DisplayName):" -ForegroundColor Cyan
                Write-Host "================================================" -ForegroundColor Cyan
                
                $methodList = @{}
                $i = 1
                $methodCount = 0

                foreach ($method in $authMethods) {
                    $methodType = $method.AdditionalProperties["@odata.type"]
                    $displayName = switch ($methodType) {
                        "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" {
                            $methodCount++
                            $deviceName = $method.AdditionalProperties["displayName"]
                            "Microsoft Authenticator: $deviceName"
                        }
                        "#microsoft.graph.phoneAuthenticationMethod" {
                            $methodCount++
                            $phoneNumber = $method.AdditionalProperties["phoneNumber"]
                            $phoneType = $method.AdditionalProperties["phoneType"]
                            "Phone ($phoneType): $phoneNumber"
                        }
                        "#microsoft.graph.emailAuthenticationMethod" {
                            $methodCount++
                            $emailAddress = $method.AdditionalProperties["emailAddress"]
                            "Email: $emailAddress"
                        }
                        "#microsoft.graph.fido2AuthenticationMethod" {
                            $methodCount++
                            $model = $method.AdditionalProperties["model"]
                            if ($model) {
                                "FIDO2 Security Key: $model"
                            } else {
                                "FIDO2 Security Key"
                            }
                        }
                        "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" {
                            $methodCount++
                            $deviceName = $method.AdditionalProperties["displayName"]
                            if ($deviceName) {
                                "Windows Hello for Business: $deviceName"
                            } else {
                                "Windows Hello for Business"
                            }
                        }
                        default { continue }
                    }

                    if ($methodType -ne "#microsoft.graph.passwordAuthenticationMethod") {
                        Write-Host "$i. $displayName" -ForegroundColor Gray
                        $methodList[$i] = @{
                            Id = $method.Id
                            Type = $methodType
                            DisplayName = $displayName
                        }
                        $i++
                    }
                }
                Write-Host "================================================" -ForegroundColor Cyan

                if ($methodCount -eq 0) {
                    Write-Host "`nNo MFA methods configured for this user." -ForegroundColor Yellow
                    $proceed = Read-Host "Press Enter to continue"
                    return
                }

                Write-Host "`nOptions:" -ForegroundColor Cyan
                Write-Host "1. Remove specific authentication method" -ForegroundColor Yellow
                Write-Host "2. Remove all authentication methods" -ForegroundColor Yellow
                Write-Host "3. Back to main menu" -ForegroundColor Yellow
                
                $choice = Read-Host "`nEnter your choice (1-3)"

                switch ($choice) {
                    "1" {
                        $methodChoice = Read-Host "`nEnter the number of the authentication method to remove (1-$($methodList.Count))"
                        if ($methodList.ContainsKey([int]$methodChoice)) {
                            $selectedMethod = $methodList[[int]$methodChoice]
                            try {
                                Remove-MgUserAuthenticationMethod -UserId $user.Id -AuthenticationMethodId $selectedMethod.Id -ErrorAction Stop
                                Write-Host "`nSuccessfully removed: $($selectedMethod.DisplayName)" -ForegroundColor Green
                            }
                            catch {
                                Write-Host "Error removing authentication method: $($_.Exception.Message)" -ForegroundColor Red
                                Write-ErrorLog -ErrorMessage "Failed to remove authentication method" -ErrorDetails $_.Exception.Message
                            }
                        }
                        else {
                            Write-Host "Invalid selection." -ForegroundColor Yellow
                        }
                    }
                    "2" {
                        $confirm = Read-Host "`nAre you sure you want to remove ALL authentication methods? (Y/N)"
                        if ($confirm -eq 'Y' -or $confirm -eq 'y') {
                            foreach ($method in $methodList.Values) {
                                try {
                                    Remove-MgUserAuthenticationMethod -UserId $user.Id -AuthenticationMethodId $method.Id -ErrorAction Stop
                                    Write-Host "Removed: $($method.DisplayName)" -ForegroundColor Green
                                }
                                catch {
                                    Write-Host "Error removing $($method.DisplayName): $($_.Exception.Message)" -ForegroundColor Red
                                    Write-ErrorLog -ErrorMessage "Failed to remove authentication method" -ErrorDetails $_.Exception.Message
                                }
                            }
                        }
                        else {
                            Write-Host "Operation cancelled." -ForegroundColor Yellow
                        }
                    }
                    "3" {
                        return
                    }
                    default {
                        Write-Host "Invalid choice. Please enter a number between 1 and 3." -ForegroundColor Yellow
                    }
                }

                $continue = Read-Host "`nDo you want to continue managing MFA methods for this user? (Y/N)"
            } while ($continue -eq 'Y' -or $continue -eq 'y')
        }
        catch {
            Write-Host "Error accessing authentication methods: $($_.Exception.Message)" -ForegroundColor Red
            Write-ErrorLog -ErrorMessage "Failed to access authentication methods" -ErrorDetails $_.Exception.Message
        }
    }
    catch {
        Write-Host "An unexpected error occurred: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Unexpected error in MFA management" -ErrorDetails $_.Exception.Message
    }
    
    $continue = Read-Host "`nPress Enter to continue"
}

# Main Script Execution
Clear-Host
Show-Banner
$maxRetries = 3
$retryCount = 0
$connected = $false

do {
    if (-not $connected) {
        $retryCount++
        if ($retryCount -gt $maxRetries) {
            Write-Host "Failed to establish Microsoft Graph connection after $maxRetries attempts. Please check your permissions and try again later." -ForegroundColor Red
            break
        }
        
        $connected = Connect-EntraGraph
        if (-not $connected) {
            Write-Host "Retrying connection... (Attempt $retryCount of $maxRetries)" -ForegroundColor Yellow
            Start-Sleep -Seconds 5
            continue
        }
    }

    Show-Menu
    $choice = Read-Host "Enter your choice"

    switch ($choice) {
        "1" { 
            try {
                List-AllUsers 
            }
            catch {
                Write-Host "Error listing users: $($_.Exception.Message)" -ForegroundColor Red
                Write-ErrorLog -ErrorMessage "Failed to list users" -ErrorDetails $_.Exception.Message
                if ($_.Exception.Message -match "Unauthorized|Authentication|Token") {
                    $connected = $false
                }
            }
        }
        "2" { 
            try {
                Add-NewUser 
            }
            catch {
                Write-Host "Error adding new user: $($_.Exception.Message)" -ForegroundColor Red
                Write-ErrorLog -ErrorMessage "Failed to add new user" -ErrorDetails $_.Exception.Message
                if ($_.Exception.Message -match "Unauthorized|Authentication|Token") {
                    $connected = $false
                }
            }
        }
        "3" { 
            try {
                Update-User 
            }
            catch {
                Write-Host "Error updating user: $($_.Exception.Message)" -ForegroundColor Red
                Write-ErrorLog -ErrorMessage "Failed to update user" -ErrorDetails $_.Exception.Message
                if ($_.Exception.Message -match "Unauthorized|Authentication|Token") {
                    $connected = $false
                }
            }
        }
        "4" { 
            try {
                Show-LicenseMenu 
            }
            catch {
                Write-Host "Error in license management: $($_.Exception.Message)" -ForegroundColor Red
                Write-ErrorLog -ErrorMessage "Failed in license management" -ErrorDetails $_.Exception.Message
                if ($_.Exception.Message -match "Unauthorized|Authentication|Token") {
                    $connected = $false
                }
            }
        }
        "5" { 
            try {
                Manage-MFA 
            }
            catch {
                Write-Host "Error in MFA management: $($_.Exception.Message)" -ForegroundColor Red
                Write-ErrorLog -ErrorMessage "Failed in MFA management" -ErrorDetails $_.Exception.Message
                if ($_.Exception.Message -match "Unauthorized|Authentication|Token") {
                    $connected = $false
                }
            }
        }
        "6" { 
            Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
            try {
                Disconnect-MgGraph
                Write-Host "Successfully disconnected from Microsoft Graph." -ForegroundColor Green
            }
            catch {
                Write-Host "Error disconnecting from Microsoft Graph: $($_.Exception.Message)" -ForegroundColor Yellow
            }
            Write-Host "Exiting script. Goodbye!" -ForegroundColor Green
            break
        }
        default { 
            Write-Host "Invalid choice. Please enter a number between 1 and 6." -ForegroundColor Yellow 
        }
    }

    if (-not $connected) {
        Write-Host "`nLost connection to Microsoft Graph. Attempting to reconnect..." -ForegroundColor Yellow
    }

} while ($choice -ne "6")
