# IT Helpdesk AD Lookup Tool

A PowerShell-based Active Directory user lookup application for IT Helpdesk staff. This tool provides a user-friendly GUI for searching and viewing Active Directory user information, making it easier for helpdesk staff to support users.

## Features

- Search for users by various attributes (username, first name, last name, display name, email)
- View detailed user information including account status, password expiration, and group memberships
- Easy-to-use interface designed for helpdesk operations

## Requirements

- Windows operating system with PowerShell 5.1 or newer
- Active Directory module for PowerShell (RSAT tools)
- Appropriate AD read permissions

## Installation

1. Clone this repository or download the files
2. Ensure you have the Active Directory PowerShell module installed
3. Run the script: `./src/IT-Helpdesk-User-Lookup.ps1`

## Usage

1. Enter a search term in the search box
2. Select the attribute to search by from the dropdown
3. Click "Search" or press Enter
4. Select a user from the results to view detailed information

## Project Structure

- `/src` - Source code
- `/docs` - Documentation
- `/tests` - Test scripts
- `/modules` - Additional PowerShell modules

## Contributing

Contributions are welcome! Please check the issues list for planned enhancements and bug fixes.

## License

MIT License