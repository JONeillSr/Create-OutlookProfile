# Create-OutlookProfile
A PowerShell script that automates Outlook profile creation for Microsoft 365

Key features:

Registry-based profile creation - Creates proper Outlook profile structure
Microsoft 365 configuration - Pre-configures Exchange Online settings
Batch processing - Handles multiple users from CSV
Error handling - Reports success/failure for each profile
Safety checks - Warns if Outlook is running and checks for existing profiles

Important considerations:

Run as Administrator - Registry modifications may require elevated privileges
Close Outlook first - Outlook should be closed during profile creation
Modern Authentication - Users will need to complete OAuth sign-in on first use
Outlook version - Script targets Office 365/2019/2021 (registry path uses 16.0)
Testing recommended - Test with a single user first before batch processing

The script creates the registry structure that Outlook expects for Microsoft 365 connections, but users will still need to complete the modern authentication flow when they first open Outlook with the new profile.

# Basic usage
.\Create-OutlookProfile.ps1 -CsvPath "C:\path\to\users.csv"

# With custom profile name
.\Create-OutlookProfile.ps1 -CsvPath "C:\path\to\users.csv" -ProfileName "Company M365"

# Set first profile as default
.\Create-OutlookProfile.ps1 -CsvPath "C:\path\to\users.csv" -SetAsDefault