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