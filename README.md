# Interactive Advanced Entra ID User Management Script

## Overview

The **Interactive Advanced Entra ID User Management Script** provides a comprehensive and user-friendly solution for managing users, licenses, groups, roles, and MFA settings in **Entra ID (formerly Azure Active Directory)**. This PowerShell-based tool simplifies administrative tasks through an intuitive interactive menu system.

## Features

- **User Management**
  - List all users with detailed properties
  - Add new users with customizable settings
  - Update existing user properties
  - Search users by name, email, or department

- **License Management**
  - View available licenses and their counts
  - Assign licenses to individual or multiple users
  - Remove licenses with confirmation
  - View current license assignments

- **Role Management**
  - Assign directory roles during user creation
  - View available roles and their IDs
  - Manage role assignments

- **MFA Management**
  - View current authentication methods
  - Add/Remove authentication methods
  - Manage FIDO2 security keys
  - Configure Windows Hello for Business

## Prerequisites

### Required Modules
```powershell
Install-Module -Name Microsoft.Graph.Authentication
Install-Module -Name Microsoft.Graph.Users
Install-Module -Name Microsoft.Graph.Identity.DirectoryManagement
Install-Module -Name Microsoft.Graph.Users.Actions
Install-Module -Name Microsoft.Graph.Identity.SignIns
```

### Required Permissions
The account running the script needs these Microsoft Graph permissions:
- `User.ReadWrite.All`
- `Group.ReadWrite.All`
- `Directory.AccessAsUser.All`
- `UserAuthenticationMethod.ReadWrite.All`

## Installation

1. Clone or download this repository
2. Ensure you have PowerShell 5.1 or later
3. Install required modules as mentioned above
4. Run the script from PowerShell ISE or PowerShell console

## Usage

### Starting the Script
```powershell
.\Interactive-EntraIDUserManagement.ps1
```

### Main Menu Options
1. **List all users**
   - View comprehensive list of all users
   - Displays key properties like name, email, and status

2. **Add a new user**
   - Enter user details (name, email, etc.)
   - Set temporary password
   - Assign licenses (optional)
   - Assign roles (optional)

3. **Update a user**
   - Modify display name
   - Update job title
   - Change department
   - Enable/disable account

4. **License Management**
   - View available licenses
   - Assign licenses to users
   - Remove licenses
   - Bulk license operations

5. **MFA Management**
   - View authentication methods
   - Add/remove methods
   - Configure security settings

### Example Operations

#### Adding a New User
```powershell
1. Select "Add a new user" from main menu
2. Enter required information:
   - Display Name
   - User Principal Name (email)
   - Mail Nickname
   - Temporary Password
3. Choose license (optional)
4. Assign directory role (optional)
```

#### Managing Licenses
```powershell
1. Select "License Management"
2. Choose user(s)
3. Select from options:
   - Add License
   - Remove License
   - View Current Licenses
```

## Error Handling

- The script includes comprehensive error handling
- All operations are logged
- User-friendly error messages
- Automatic retry for certain operations

## Security Features

- Password complexity requirements
- Forced password change at first login
- MFA configuration options
- Secure credential handling

## Author
- **Name**: Mezba Uddin
- **Version**: 2.2
- **Last Updated**: 2024-11-06

## Support

For issues, questions, or contributions:
1. Open an issue in the GitHub repository
2. Provide detailed information about the problem
3. Include error messages and script version

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Microsoft Graph PowerShell SDK team
- PowerShell community
- All contributors and testers
