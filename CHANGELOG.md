# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added
- User actions menu with the following features:
  - Unlock account functionality
  - Reset password with option to require password change at next logon
  - Enable/disable account capability
  - Quick refresh of user details
- Export functionality for user details (CSV and HTML formats)
- Visual indicators for account status (color coding)
  - Red for disabled or locked accounts
  - Orange for warnings (soon-to-expire passwords)
  - Green for good status
  - Blue for special configurations
- Color legend in user details view

### Changed
- Improved search functionality to search across multiple fields simultaneously (username, first name, last name, display name, email)
- Removed search field dropdown for a more streamlined interface
- Enhanced UI layout for better usability
- Increased details box size to display more user information
- Improved keyboard support with Ctrl+A (Select All) and standard copy/paste operations
- Enhanced error handling and null checking for more robust operation

### Initial Features
- Basic user search functionality
- Detailed user information display
- Group membership listing

## [0.1.0] - 2025-03-14

### Added
- Initial repository setup
- Base PowerShell script for AD user lookup