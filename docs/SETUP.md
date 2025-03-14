# Setup Guide

## Prerequisites

### Required Software

1. Windows PowerShell 5.1 or newer
2. Active Directory module for PowerShell

### Installing the AD Module

If you don't have the Active Directory module installed:

1. Open PowerShell as Administrator
2. Run the following command to install the RSAT AD tools:
   ```powershell
   Add-WindowsCapability -Online -Name Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0
   ```
3. Verify installation:
   ```powershell
   Import-Module ActiveDirectory
   Get-Module ActiveDirectory
   ```

## Application Setup

1. Clone this repository:
   ```
   git clone https://github.com/Zolkiev/IT-Helpdesk-AD-Lookup-Tool.git
   ```

2. Navigate to the source directory:
   ```
   cd IT-Helpdesk-AD-Lookup-Tool/src
   ```

3. Run the application:
   ```powershell
   ./IT-Helpdesk-User-Lookup.ps1
   ```

## Required Permissions

The user running this script needs at minimum:

- Read access to the Active Directory domain
- For advanced features (unlocking accounts, resetting passwords), additional privileges are required

## Using the Application

### Basic Search
1. Enter a search term in the search box
2. Click "Search" or press Enter
3. The tool will search across multiple fields (username, first name, last name, display name, email)
4. Select a user from the results to view detailed information

### Exporting User Details
1. Search for and select a user from the results
2. Click the "Export User Details" button
3. Choose either "Export to CSV" or "Export to HTML"
4. Select a location to save the file
5. The exported file will contain all user details including:
   - Basic information (name, email, etc.)
   - Account status
   - Login information
   - Group memberships

## Troubleshooting

### Common Issues

1. **Module not found error**:
   Ensure the Active Directory module is installed properly.

2. **Access denied errors**:
   Check that your account has sufficient permissions to query AD.

3. **Script execution policy error**:
   You may need to adjust the PowerShell execution policy:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

4. **Export feature issues**:
   If you encounter problems with the export functionality, ensure you have write permissions to the location where you're trying to save the files.