# Interactive Advanced Entra ID User Management Script

## Overview

This **Interactive Advanced Entra ID User Management Script** provides a robust and user-friendly way to manage users, licenses, groups, roles, and MFA settings in **Entra ID (formerly Azure AD)**. The interactive menu allows administrators to perform complex management tasks with ease.

## Features

- **List Users**: Retrieve detailed information about users, including names, email addresses, and more.
- **Add Users**: Add new users with options to assign licenses and roles.
- **Update Users**: Modify user properties such as display name, job title, and group memberships.
- **Manage Licenses**: Assign or remove licenses for single or multiple users with real-time selection.
- **Manage MFA**: View, add, and remove authentication methods for users.
- **Search Users**: Search users dynamically using names, emails, or departments.

## Requirements

- **Microsoft Graph Module**: Install using:
  ```powershell
  Install-Module -Name Microsoft.Graph -Scope CurrentUser
Permissions: The script requires scopes such as:
User.ReadWrite.All
Group.ReadWrite.All
Directory.AccessAsUser.All
UserAuthenticationMethod.ReadWrite.All
Authentication: Use Connect-MgGraph to authenticate to Microsoft Graph before executing operations.
Usage
Step 1: Install Dependencies
Ensure you have the necessary permissions and modules installed:

powershell
Copy code
Install-Module Microsoft.Graph -Scope CurrentUser
Step 2: Run the Script
Authenticate to Microsoft Graph when prompted.
Use the interactive menu to choose and execute tasks.
Step 3: Menu Options
List Users: Fetch a list of all users with detailed properties.
Add New User: Add a user with licenses and roles.
Update User: Modify user attributes or manage group memberships.
License Management:
Assign or remove licenses for single or multiple users.
MFA Management: Manage multi-factor authentication methods for users.
Exit: Disconnect from Microsoft Graph and exit.
Example Output
markdown
Copy code
==========================
Entra ID User Management
==========================
1. List all users
2. Add a new user
3. Update a user
4. License Management
5. MFA Management
6. Exit
==========================
Enter your choice:
Author
Mezba Uddin

Version
2.2
Last Updated: 2024-11-06
License
This project is licensed under the MIT License.

Contributions
Contributions and feedback are welcome! Submit pull requests or issues via GitHub.
