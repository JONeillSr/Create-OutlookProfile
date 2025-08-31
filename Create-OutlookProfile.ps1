<#
.SYNOPSIS
    Creates new Outlook profiles for Microsoft 365 from a CSV file containing user UPNs.

.DESCRIPTION
    This script automates the creation of Outlook profiles configured for Microsoft 365 (Exchange Online).
    It reads User Principal Names (UPNs) from a CSV file and creates corresponding Outlook profiles
    with the necessary registry entries for Exchange Online connectivity.
    
    The script creates registry entries in the appropriate Outlook and Windows Messaging Subsystem
    locations to establish the profile structure. Users will still need to complete Modern Authentication
    (OAuth) when first accessing their mailbox.

.PARAMETER CsvPath
    Mandatory. The full path to the CSV file containing user UPNs. The CSV must have a column named 'UPN'
    with email addresses for each user.

.PARAMETER ProfileName
    Optional. The base name for the Outlook profiles. Default is "Microsoft 365 Profile".
    Each profile will be created as "ProfileName - UPN" to ensure uniqueness.

.PARAMETER SetAsDefault
    Optional switch. If specified, the first profile created will be set as the default Outlook profile.
    Only affects the first profile when processing multiple users.

.INPUTS
    CSV file with UPN column containing email addresses.

.OUTPUTS
    Registry entries for Outlook profiles and console output showing creation status.

.EXAMPLE
    .\Create-OutlookProfile.ps1 -CsvPath "C:\Users\Admin\users.csv"
    
    Creates Outlook profiles for all users in the CSV file using the default profile name.

.EXAMPLE
    .\Create-OutlookProfile.ps1 -CsvPath "C:\temp\employees.csv" -ProfileName "Company Email"
    
    Creates profiles with the base name "Company Email" for all users in the CSV file.

.EXAMPLE
    .\Create-OutlookProfile.ps1 -CsvPath "C:\data\users.csv" -SetAsDefault
    
    Creates profiles and sets the first one as the default Outlook profile.

.NOTES
    Author: John O'Neill Sr.
    Company: Azure Innovators
    Date: 08/28/2025
    Version: 2.0
    Change Purpose: Converted from COM automation to Microsoft Graph API

    Prerequisites:
                    PowerShell 5.1 or later
                    Administrative privileges recommended
    
    IMPORTANT REQUIREMENTS:
    - Outlook must be closed before running this script
    - Run with Administrative privileges for registry access
    - CSV file must contain 'UPN' column with valid email addresses
    - Targets Office 365/2019/2021 (registry path 16.0)
    - Users must complete Modern Authentication on first Outlook launch
    
    LIMITATIONS:
    - Does not handle complex Exchange configurations
    - Profiles created are basic Exchange Online configurations
    - Does not migrate existing data or settings
    - Registry modifications may require restart of Outlook

.LINK
    https://docs.microsoft.com/en-us/outlook/
    
.LINK
    https://docs.microsoft.com/en-us/microsoft-365/admin/email/
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$false)]
    [string]$ProfileName = "M365 Profile",
    
    [Parameter(Mandatory=$false)]
    [switch]$SetAsDefault
)

# Function to create Outlook profile
function New-OutlookProfile {
    param(
        [string]$ProfileName,
        [string]$EmailAddress,
        [bool]$SetAsDefault = $false
    )
    
    try {
        Write-Host "Creating Outlook profile: $ProfileName for $EmailAddress" -ForegroundColor Green
        
        # Registry paths for Outlook profiles
        $outlookProfilesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
        $windowsMailPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
        
        # Check if profile already exists
        if (Test-Path "$outlookProfilesPath\$ProfileName") {
            Write-Warning "Profile '$ProfileName' already exists. Skipping creation."
            return $false
        }
        
        # Create the profile registry structure
        New-Item -Path "$outlookProfilesPath\$ProfileName" -Force | Out-Null
        New-Item -Path "$windowsMailPath\$ProfileName" -Force | Out-Null
        
        # Set profile properties
        Set-ItemProperty -Path "$outlookProfilesPath\$ProfileName" -Name "NextAccountID" -Value 1 -Type DWord
        Set-ItemProperty -Path "$outlookProfilesPath\$ProfileName" -Name "NextServiceID" -Value 2 -Type DWord
        
        # Create account registry entries
        $accountPath = "$outlookProfilesPath\$ProfileName\9375CFF0413111d3B88A00104B2A6676"
        New-Item -Path $accountPath -Force | Out-Null
        
        # Configure Exchange account settings
        $exchangePath = "$accountPath\00000001"
        New-Item -Path $exchangePath -Force | Out-Null
        
        # Set Exchange account properties
        Set-ItemProperty -Path $exchangePath -Name "Account Name" -Value $EmailAddress -Type String
        Set-ItemProperty -Path $exchangePath -Name "Display Name" -Value $EmailAddress -Type String
        Set-ItemProperty -Path $exchangePath -Name "Email" -Value $EmailAddress -Type String
        Set-ItemProperty -Path $exchangePath -Name "Exchange Server" -Value "outlook.office365.com" -Type String
        Set-ItemProperty -Path $exchangePath -Name "User" -Value $EmailAddress -Type String
        
        # Set service provider GUID for Exchange
        Set-ItemProperty -Path $exchangePath -Name "Service UID" -Value ([byte[]](0x13,0x1F,0xD8,0xAB,0x86,0x8B,0xCE,0x11,0x9F,0xCB,0x00,0xAA,0x00,0x6C,0x45,0x4C)) -Type Binary
        
        # Configure Windows Messaging Subsystem profile
        $wmsProfilePath = "$windowsMailPath\$ProfileName"
        Set-ItemProperty -Path $wmsProfilePath -Name "NextServiceID" -Value 2 -Type DWord
        
        # Add service to WMS profile
        $wmsServicePath = "$wmsProfilePath\13dbb0c8aa05101a9bb000aa002fc45a"
        New-Item -Path $wmsServicePath -Force | Out-Null
        Set-ItemProperty -Path $wmsServicePath -Name "Service Name" -Value "MSEMS" -Type String
        
        # Set as default profile if requested
        if ($SetAsDefault) {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name "DefaultProfile" -Value $ProfileName -Type String
            Write-Host "Set '$ProfileName' as default profile" -ForegroundColor Yellow
        }
        
        Write-Host "Successfully created profile: $ProfileName" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Error "Failed to create profile '$ProfileName': $($_.Exception.Message)"
        return $false
    }
}

# Function to read CSV and process users
function Process-UsersFromCsv {
    param(
        [string]$CsvPath,
        [string]$BaseProfileName,
        [bool]$SetAsDefault
    )
    
    if (-not (Test-Path $CsvPath)) {
        Write-Error "CSV file not found: $CsvPath"
        return
    }
    
    try {
        $users = Import-Csv $CsvPath
        $successCount = 0
        $failCount = 0
        
        foreach ($user in $users) {
            # Assuming CSV has 'UPN' column, adjust as needed
            $upn = $user.UPN
            if (-not $upn) {
                Write-Warning "No UPN found in row. Skipping."
                $failCount++
                continue
            }
            
            # Create unique profile name
            $profileName = "$BaseProfileName - $upn"
            
            # Create the profile
            if (New-OutlookProfile -ProfileName $profileName -EmailAddress $upn -SetAsDefault $SetAsDefault) {
                $successCount++
            } else {
                $failCount++
            }
            
            # Only set the first profile as default if requested
            $SetAsDefault = $false
        }
        
        Write-Host "`nProfile Creation Summary:" -ForegroundColor Cyan
        Write-Host "Successful: $successCount" -ForegroundColor Green
        Write-Host "Failed: $failCount" -ForegroundColor Red
        
    } catch {
        Write-Error "Error processing CSV file: $($_.Exception.Message)"
    }
}

# Main execution
Write-Host "Outlook Microsoft 365 Profile Creator" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan

# Check if Outlook is running
$outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
if ($outlookProcess) {
    Write-Warning "Outlook is currently running. Please close Outlook before creating profiles."
    $response = Read-Host "Do you want to continue anyway? (y/N)"
    if ($response -ne 'y' -and $response -ne 'Y') {
        Write-Host "Script cancelled." -ForegroundColor Yellow
        exit
    }
}

# Process the CSV file
Process-UsersFromCsv -CsvPath $CsvPath -BaseProfileName $ProfileName -SetAsDefault $SetAsDefault.IsPresent

Write-Host "`nScript completed. Please restart Outlook to see the new profiles." -ForegroundColor Green
Write-Host "Note: Users may need to complete Modern Authentication when first accessing their mailbox." -ForegroundColor Yellow
param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    
    [Parameter(Mandatory=$false)]
    [string]$ProfileName = "Microsoft 365 Profile",
    
    [Parameter(Mandatory=$false)]
    [switch]$SetAsDefault
)

# Function to create Outlook profile
function New-OutlookProfile {
    param(
        [string]$ProfileName,
        [string]$EmailAddress,
        [bool]$SetAsDefault = $false
    )
    
    try {
        Write-Host "Creating Outlook profile: $ProfileName for $EmailAddress" -ForegroundColor Green
        
        # Registry paths for Outlook profiles
        $outlookProfilesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"
        $windowsMailPath = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles"
        
        # Check if profile already exists
        if (Test-Path "$outlookProfilesPath\$ProfileName") {
            Write-Warning "Profile '$ProfileName' already exists. Skipping creation."
            return $false
        }
        
        # Create the profile registry structure
        New-Item -Path "$outlookProfilesPath\$ProfileName" -Force | Out-Null
        New-Item -Path "$windowsMailPath\$ProfileName" -Force | Out-Null
        
        # Set profile properties
        Set-ItemProperty -Path "$outlookProfilesPath\$ProfileName" -Name "NextAccountID" -Value 1 -Type DWord
        Set-ItemProperty -Path "$outlookProfilesPath\$ProfileName" -Name "NextServiceID" -Value 2 -Type DWord
        
        # Create account registry entries
        $accountPath = "$outlookProfilesPath\$ProfileName\9375CFF0413111d3B88A00104B2A6676"
        New-Item -Path $accountPath -Force | Out-Null
        
        # Configure Exchange account settings
        $exchangePath = "$accountPath\00000001"
        New-Item -Path $exchangePath -Force | Out-Null
        
        # Set Exchange account properties
        Set-ItemProperty -Path $exchangePath -Name "Account Name" -Value $EmailAddress -Type String
        Set-ItemProperty -Path $exchangePath -Name "Display Name" -Value $EmailAddress -Type String
        Set-ItemProperty -Path $exchangePath -Name "Email" -Value $EmailAddress -Type String
        Set-ItemProperty -Path $exchangePath -Name "Exchange Server" -Value "outlook.office365.com" -Type String
        Set-ItemProperty -Path $exchangePath -Name "User" -Value $EmailAddress -Type String
        
        # Set service provider GUID for Exchange
        Set-ItemProperty -Path $exchangePath -Name "Service UID" -Value ([byte[]](0x13,0x1F,0xD8,0xAB,0x86,0x8B,0xCE,0x11,0x9F,0xCB,0x00,0xAA,0x00,0x6C,0x45,0x4C)) -Type Binary
        
        # Configure Windows Messaging Subsystem profile
        $wmsProfilePath = "$windowsMailPath\$ProfileName"
        Set-ItemProperty -Path $wmsProfilePath -Name "NextServiceID" -Value 2 -Type DWord
        
        # Add service to WMS profile
        $wmsServicePath = "$wmsProfilePath\13dbb0c8aa05101a9bb000aa002fc45a"
        New-Item -Path $wmsServicePath -Force | Out-Null
        Set-ItemProperty -Path $wmsServicePath -Name "Service Name" -Value "MSEMS" -Type String
        
        # Set as default profile if requested
        if ($SetAsDefault) {
            Set-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name "DefaultProfile" -Value $ProfileName -Type String
            Write-Host "Set '$ProfileName' as default profile" -ForegroundColor Yellow
        }
        
        Write-Host "Successfully created profile: $ProfileName" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Error "Failed to create profile '$ProfileName': $($_.Exception.Message)"
        return $false
    }
}

# Function to read CSV and process users
function Process-UsersFromCsv {
    param(
        [string]$CsvPath,
        [string]$BaseProfileName,
        [bool]$SetAsDefault
    )
    
    if (-not (Test-Path $CsvPath)) {
        Write-Error "CSV file not found: $CsvPath"
        return
    }
    
    try {
        $users = Import-Csv $CsvPath
        $successCount = 0
        $failCount = 0
        
        foreach ($user in $users) {
            # Assuming CSV has 'UPN' column, adjust as needed
            $upn = $user.UPN
            if (-not $upn) {
                Write-Warning "No UPN found in row. Skipping."
                $failCount++
                continue
            }
            
            # Create unique profile name
            $profileName = "$BaseProfileName - $upn"
            
            # Create the profile
            if (New-OutlookProfile -ProfileName $profileName -EmailAddress $upn -SetAsDefault $SetAsDefault) {
                $successCount++
            } else {
                $failCount++
            }
            
            # Only set the first profile as default if requested
            $SetAsDefault = $false
        }
        
        Write-Host "`nProfile Creation Summary:" -ForegroundColor Cyan
        Write-Host "Successful: $successCount" -ForegroundColor Green
        Write-Host "Failed: $failCount" -ForegroundColor Red
        
    } catch {
        Write-Error "Error processing CSV file: $($_.Exception.Message)"
    }
}

# Main execution
Write-Host "Outlook Microsoft 365 Profile Creator" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan

# Check if Outlook is running
$outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
if ($outlookProcess) {
    Write-Warning "Outlook is currently running. Please close Outlook before creating profiles."
    $response = Read-Host "Do you want to continue anyway? (y/N)"
    if ($response -ne 'y' -and $response -ne 'Y') {
        Write-Host "Script cancelled." -ForegroundColor Yellow
        exit
    }
}

# Process the CSV file
Process-UsersFromCsv -CsvPath $CsvPath -BaseProfileName $ProfileName -SetAsDefault $SetAsDefault.IsPresent

Write-Host "`nScript completed. Please restart Outlook to see the new profiles." -ForegroundColor Green
Write-Host "Note: Users may need to complete Modern Authentication when first accessing their mailbox." -ForegroundColor Yellow