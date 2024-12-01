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

# Import required Microsoft.Graph modules
$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Users",
    "Microsoft.Graph.Identity.DirectoryManagement",
    "Microsoft.Graph.Users.Actions",
    "Microsoft.Graph.Identity.SignIns"
)

foreach ($module in $requiredModules) {
    if (!(Get-Module -Name $module -ListAvailable)) {
        Write-Host "Installing $module..." -ForegroundColor Cyan
        Install-Module -Name $module -Force -Scope CurrentUser
    }
    Import-Module -Name $module -Force
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

# Function to display main menu
function Show-Menu {
    Write-Host "==========================" -ForegroundColor Green
    Write-Host "Entra ID User Management" -ForegroundColor Cyan
    Write-Host "==========================" -ForegroundColor Green
    Write-Host "1. List all users" -ForegroundColor Cyan
    Write-Host "2. Add a new user" -ForegroundColor Cyan
    Write-Host "3. Update a user" -ForegroundColor Cyan
    Write-Host "4. License Management" -ForegroundColor Cyan
    Write-Host "5. MFA Management" -ForegroundColor Cyan
    Write-Host "6. Exit" -ForegroundColor Cyan
    Write-Host "==========================" -ForegroundColor Green
}

# Function to show license management menu
function Show-LicenseMenu {
    Write-Host "`n==========================" -ForegroundColor Green
    Write-Host "License Management" -ForegroundColor Cyan
    Write-Host "==========================" -ForegroundColor Green
    Write-Host "1. Manage Single User" -ForegroundColor Cyan
    Write-Host "2. Manage Multiple Users" -ForegroundColor Cyan
    Write-Host "3. Back to Main Menu" -ForegroundColor Cyan
    Write-Host "==========================" -ForegroundColor Green

    $choice = Read-Host "`nEnter your choice (1-3)"
    
    switch ($choice) {
        "1" { Manage-Licenses }
        "2" { Manage-BulkLicenses }
        "3" { return }
        default { 
            Write-Host "Invalid choice. Please try again." -ForegroundColor Yellow
            Show-LicenseMenu
        }
    }
}

# Function to list all users
function List-AllUsers {
    try {
        Write-Host "Fetching all Entra ID users..." -ForegroundColor Cyan
        $users = Get-MgUser -All -ErrorAction Stop
        
        if ($users.Count -eq 0) {
            Write-Host "No users found in the directory." -ForegroundColor Yellow
            return
        }

        Write-Host "`nUser List:" -ForegroundColor Green
        Write-Host "==========" -ForegroundColor Green
        foreach ($user in $users) {
            Write-Host "Display Name: $($user.DisplayName)" -ForegroundColor Cyan
            Write-Host "Email: $($user.Mail)"
            Write-Host "ID: $($user.Id)"
            Write-Host "Account Enabled: $($user.AccountEnabled)"
            Write-Host "----------"
        }
    }
    catch {
        Write-Host "Error fetching users: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Failed to fetch users" -ErrorDetails $_.Exception.Message
    }
    
    $continue = Read-Host "`nPress Enter to continue"
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
        $script:licenseOptions = @{}
        
        Write-Host "`nAvailable Licenses:"
        Write-Host "==================="
        foreach ($sub in $subscriptions) {
            $availableUnits = $sub.PrepaidUnits.Enabled - $sub.ConsumedUnits
            Write-Host "$i. $($sub.SkuPartNumber) - Available: $availableUnits" -ForegroundColor Cyan
            $script:licenseOptions[$i.ToString()] = @{
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
            
            if ($script:licenseOptions[$selection]) {
                if ($script:licenseOptions[$selection].AvailableUnits -le 0) {
                    Write-Host "No available licenses for $($script:licenseOptions[$selection].SkuPartNumber)" -ForegroundColor Yellow
                    continue
                }
                return $script:licenseOptions[$selection].SkuId
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
            Write-Host "- $($license.SkuPartNumber)" -ForegroundColor Cyan
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
        Write-Host "`nAdding New User" -ForegroundColor Cyan
        Write-Host "==============" -ForegroundColor Cyan
        
        # Collect user information with validation
        do {
            $DisplayName = Read-Host "Enter the display name for the new user"
        } while ([string]::IsNullOrWhiteSpace($DisplayName))
        
        do {
            $UserPrincipalName = Read-Host "Enter the user principal name (e.g., john.doe@yourdomain.com)"
            if (-not ($UserPrincipalName -match "^[^@]+@[^@]+\.[^@]+$")) {
                Write-Host "Invalid email format. Please try again." -ForegroundColor Yellow
                continue
            }
            
            # Check if user already exists
            $existingUser = Get-MgUser -Filter "UserPrincipalName eq '$UserPrincipalName'" -ErrorAction SilentlyContinue
            if ($existingUser) {
                Write-Host "A user with this email already exists. Please try a different email." -ForegroundColor Yellow
                continue
            }
            break
        } while ($true)
        
        do {
            $MailNickname = Read-Host "Enter the mail nickname for the user (will be used in email address)"
            if ([string]::IsNullOrWhiteSpace($MailNickname)) {
                Write-Host "Mail nickname cannot be empty. Please try again." -ForegroundColor Yellow
                continue
            }
            break
        } while ($true)
        
        do {
            $Password = Read-Host "Enter a temporary password for the user (minimum 8 characters, must include uppercase, lowercase, number, and special character)"
            if ($Password.Length -lt 8 -or 
                -not ($Password -match "[A-Z]") -or 
                -not ($Password -match "[a-z]") -or 
                -not ($Password -match "[0-9]") -or 
                -not ($Password -match "[^A-Za-z0-9]")) {
                Write-Host "Password does not meet complexity requirements. Please try again." -ForegroundColor Yellow
                continue
            }
            break
        } while ($true)

        Write-Host "`nSelect a license to assign (optional):" -ForegroundColor Cyan
        $LicenseSkuId = Get-AvailableLicenses
        
        Write-Host "`nCreating new user: $DisplayName..." -ForegroundColor Cyan
        $newUser = New-MgUser -AccountEnabled $true `
            -DisplayName $DisplayName `
            -UserPrincipalName $UserPrincipalName `
            -MailNickname $MailNickname `
            -PasswordProfile @{
                Password = $Password
                ForceChangePasswordNextSignIn = $true
                ForceChangePasswordNextSignInWithMfa = $true
            } -ErrorAction Stop
            
        Write-Host "User $DisplayName created successfully!" -ForegroundColor Green

        if ($LicenseSkuId) {
            Write-Host "Assigning license..." -ForegroundColor Cyan
            Set-UserLicense -UserId $newUser.Id -LicenseSkuId $LicenseSkuId -RemoveLicense $false
            Write-Host "License assigned successfully!" -ForegroundColor Green
        }

        # Ask about role assignment
        $assignRole = Read-Host "`nDo you want to assign a directory role to this user? (Y/N)"
        if ($assignRole -eq 'Y' -or $assignRole -eq 'y') {
            $roles = Get-MgDirectoryRole
            Write-Host "`nAvailable Directory Roles:" -ForegroundColor Cyan
            Write-Host "======================" -ForegroundColor Cyan
            foreach ($role in $roles) {
                Write-Host "$($role.Id): $($role.DisplayName)" -ForegroundColor Yellow
            }
            
            $RoleId = Read-Host "`nEnter the Directory Role ID to assign"
            if (-not [string]::IsNullOrWhiteSpace($RoleId)) {
                New-MgDirectoryRoleMember -DirectoryRoleId $RoleId -DirectoryObjectId $newUser.Id -ErrorAction Stop
                Write-Host "Role assigned successfully!" -ForegroundColor Green
            }
        }

        Write-Host "`nUser creation completed successfully!" -ForegroundColor Green
        Write-Host "Please provide these credentials to the user:" -ForegroundColor Cyan
        Write-Host "Username: $UserPrincipalName" -ForegroundColor Yellow
        Write-Host "Temporary Password: $Password" -ForegroundColor Yellow
        Write-Host "Note: The user will be required to change their password upon first login." -ForegroundColor Cyan
    }
    catch {
        Write-Host "Error creating user: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Failed to create new user" -ErrorDetails $_.Exception.Message
    }
    
    $continue = Read-Host "`nPress Enter to continue"
}

# Function to update user properties
function Update-User {
    try {
        $user = Find-EntraIDUser
        if (-not $user) { return }

        do {
            Write-Host "`nUpdate Options for $($user.DisplayName):" -ForegroundColor Cyan
            Write-Host "1. Update Display Name" -ForegroundColor Yellow
            Write-Host "2. Update Job Title" -ForegroundColor Yellow
            Write-Host "3. Update Department" -ForegroundColor Yellow
            Write-Host "4. Enable/Disable Account" -ForegroundColor Yellow
            Write-Host "5. Back to Main Menu" -ForegroundColor Yellow
            
            $choice = Read-Host "`nEnter your choice (1-5)"
            
            switch ($choice) {
                "1" {
                    $newDisplayName = Read-Host "Enter new display name"
                    Update-MgUser -UserId $user.Id -DisplayName $newDisplayName
                    Write-Host "Display name updated successfully!" -ForegroundColor Green
                }
                "2" {
                    $newJobTitle = Read-Host "Enter new job title"
                    Update-MgUser -UserId $user.Id -JobTitle $newJobTitle
                    Write-Host "Job title updated successfully!" -ForegroundColor Green
                }
                "3" {
                    $newDepartment = Read-Host "Enter new department"
                    Update-MgUser -UserId $user.Id -Department $newDepartment
                    Write-Host "Department updated successfully!" -ForegroundColor Green
                }
                "4" {
                    $currentStatus = (Get-MgUser -UserId $user.Id).AccountEnabled
                    $newStatus = -not $currentStatus
                    Update-MgUser -UserId $user.Id -AccountEnabled $newStatus
                    $statusText = if ($newStatus) { "enabled" } else { "disabled" }
                    Write-Host "Account has been $statusText!" -ForegroundColor Green
                }
                "5" { return }
                default { Write-Host "Invalid choice. Please try again." -ForegroundColor Yellow }
            }
            
            $user = Get-MgUser -UserId $user.Id
        } while ($choice -ne "5")
    }
    catch {
        Write-Host "Error updating user: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Failed to update user" -ErrorDetails $_.Exception.Message
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
                        Set-UserLicense -UserId $user.Id -LicenseSkuId $LicenseSkuId -RemoveLicense $false
                        Write-Host "License assigned successfully to $($user.UserPrincipalName)" -ForegroundColor Green
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
                            Set-UserLicense -UserId $user.Id -LicenseSkuId $userLicenses[$selection-1].SkuId -RemoveLicense $true
                            Write-Host "License removed successfully from $($user.UserPrincipalName)" -ForegroundColor Green
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
                    Set-UserLicense -UserId $user.Id -LicenseSkuId $LicenseSkuId -RemoveLicense $false
                    Write-Host "License assigned successfully to $($user.UserPrincipalName)" -ForegroundColor Green
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
                    Set-UserLicense -UserId $selected.User.Id -LicenseSkuId $selected.LicenseSkuId -RemoveLicense $true
                    Write-Host "License removed successfully from $($selected.User.UserPrincipalName)" -ForegroundColor Green
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

        # Get current MFA methods
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
                        "microsoftAuthenticator" {
                            "Microsoft Authenticator"
                        }
                        "phoneAuthentication" {
                            $phoneNumber = $method.AdditionalProperties.phoneNumber
                            "Phone Authentication: $phoneNumber"
                        }
                        "emailAuthentication" {
                            $emailAddress = $method.AdditionalProperties.emailAddress
                            "Email Authentication: $emailAddress"
                        }
                        "passwordAuthentication" {
                            "Password"
                        }
                        "windowsHelloForBusiness" {
                            $deviceName = $method.AdditionalProperties.displayName
                            "Windows Hello for Business: $deviceName"
                        }
                        "fido2SecurityKey" {
                            $model = $method.AdditionalProperties.model
                            "FIDO2 Security Key: $model"
                        }
                        "temporaryAccessPass" {
                            "Temporary Access Pass"
                        }
                        default {
                            $methodType
                        }
                    }

                    if ($methodType -ne "passwordAuthentication") {
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

# Function to assign license
function Set-UserLicense {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserId,
        [Parameter(Mandatory = $true)]
        [string]$LicenseSkuId,
        [bool]$RemoveLicense = $false
    )
    
    try {
        if ($RemoveLicense) {
            # Remove license
            Set-MgUserLicense -UserId $UserId -AddLicenses @() -RemoveLicenses @($LicenseSkuId)
        } else {
            # Add license
            $addLicenseObj = @{
                SkuId = $LicenseSkuId
                DisabledPlans = @()
            }
            Set-MgUserLicense -UserId $UserId -AddLicenses @($addLicenseObj) -RemoveLicenses @()
        }
        return $true
    }
    catch {
        Write-Host "Error managing license: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Failed to manage license" -ErrorDetails $_.Exception.Message
        return $false
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

        $user = Find-EntraIDUser
        if (-not $user) { return }

        do {
            Write-Host "`nLicense Management Menu for $($user.DisplayName)" -ForegroundColor Green
            Write-Host "=====================" -ForegroundColor Green
            Write-Host "1. Add License" -ForegroundColor Cyan
            Write-Host "2. Remove License" -ForegroundColor Cyan
            Write-Host "3. View Current Licenses" -ForegroundColor Cyan
            Write-Host "4. Back to Main Menu" -ForegroundColor Cyan
            Write-Host "=====================" -ForegroundColor Green

            $choice = Read-Host "`nEnter your choice (1-4)"

            switch ($choice) {
                "1" {
                    Write-Host "`nSelect a license to assign:"
                    $LicenseSkuId = Get-AvailableLicenses
                    if ($LicenseSkuId) {
                        Set-UserLicense -UserId $user.Id -LicenseSkuId $LicenseSkuId -RemoveLicense $false
                        Write-Host "License assigned successfully!" -ForegroundColor Green
                    }
                }
                "2" {
                    $licenses = Get-MgUserLicenseDetail -UserId $user.Id
                    if ($licenses.Count -eq 0) {
                        Write-Host "User has no licenses to remove." -ForegroundColor Yellow
                        continue
                    }

                    Write-Host "`nCurrent Licenses:"
                    for ($i = 0; $i -lt $licenses.Count; $i++) {
                        Write-Host "$($i + 1). $($licenses[$i].SkuPartNumber)" -ForegroundColor Cyan
                    }

                    $selection = Read-Host "`nSelect license number to remove (1-$($licenses.Count))"
                    if ([int]::TryParse($selection, [ref]$null)) {
                        $index = [int]$selection - 1
                        if ($index -ge 0 -and $index -lt $licenses.Count) {
                            Set-UserLicense -UserId $user.Id -LicenseSkuId $licenses[$index].SkuId -RemoveLicense $true
                            Write-Host "License removed successfully!" -ForegroundColor Green
                        }
                        else {
                            Write-Host "Invalid selection." -ForegroundColor Yellow
                        }
                    }
                }
                "3" {
                    $licenses = Get-MgUserLicenseDetail -UserId $user.Id
                    if ($licenses.Count -eq 0) {
                        Write-Host "`nUser has no licenses assigned." -ForegroundColor Yellow
                    }
                    else {
                        Write-Host "`nCurrent Licenses:" -ForegroundColor Green
                        foreach ($license in $licenses) {
                            Write-Host "- $($license.SkuPartNumber)" -ForegroundColor Cyan
                        }
                    }
                    $continue = Read-Host "`nPress Enter to continue"
                }
                "4" { return }
                default { Write-Host "Invalid choice. Please try again." -ForegroundColor Yellow }
            }
        } while ($true)
    }
    catch {
        Write-Host "Error in license management: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog -ErrorMessage "Error in license management" -ErrorDetails $_.Exception.Message
        $continue = Read-Host "`nPress Enter to continue"
    }
}

# Main Script Execution
do {
    # Ensure Graph connection at script start
    if (-not (Connect-EntraGraph)) {
        Write-Host "Unable to proceed without Microsoft Graph connection. Please ensure you have the necessary permissions and try again." -ForegroundColor Red
        break
    }

    Show-Menu
    $choice = Read-Host "Enter your choice"

    switch ($choice) {
        "1" { List-AllUsers }
        "2" { Add-NewUser }
        "3" { Update-User }
        "4" { Show-LicenseMenu }
        "5" { Manage-MFA }
        "6" { 
            Write-Host "Disconnecting from Microsoft Graph..." -ForegroundColor Cyan
            Disconnect-MgGraph
            Write-Host "Exiting script. Goodbye!" -ForegroundColor Green 
        }
        default { Write-Host "Invalid choice. Please try again." -ForegroundColor Yellow }
    }
} while ($choice -ne "6")
