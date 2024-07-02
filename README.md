Overview

This PowerShell script is designed to streamline the process of updating various software applications and managing OWL firmware updates in FLEX classrooms. The script automates the update process for MS Paint, Photos, Snip Tool, Zoom Client, Aver U50 Document Scanner app, MS Office Desktop Apps, Chrome, Firefox, and checks the Okta Plugin. It also includes a function to verify network connectivity to OWL devices and placeholders for digital signing. Additionally, instructions are provided to convert the PowerShell script into an executable (.exe) file.
Features

    Update Applications: Automatically updates MS Paint, Photos, Snip Tool, Zoom Client using winget.
    Update Aver U50 Document Scanner App: Downloads and installs the latest firmware for the Aver U50 Document Scanner.
    Update MS Office Desktop Apps: Updates Microsoft Office applications (PowerPoint, Word, Excel).
    Update Browsers: Updates Google Chrome and Mozilla Firefox browsers.
    Check Okta Plugin: Verifies the installation and version of the Okta plugin in Chrome and Firefox.
    OWL Device Connectivity Check: Tests the network connectivity of OWL devices.
    Digital Signing Placeholder: Provides a placeholder for adding digital signatures to the script.
    Conversion to Executable: Instructions included to convert the PowerShell script to an executable file.

Prerequisites

    PowerShell 5.1 or later.
    winget (Windows Package Manager) installed.
    Administrator Privileges: The script may require administrator privileges to install updates.
    Internet Connection: Required for downloading updates and verifying connectivity.

Script Details
Parameters

    OwlDeviceIP: The IP address of the OWL device. Default is set to "192.168.1.100".

Functions
Test-DeviceConnectivity

Tests the network connectivity to the specified OWL device.
Update-Applications

Updates MS Paint, Photos, Snip Tool, and Zoom Client using winget.
Update-AverU50

Downloads and installs the latest firmware for the Aver U50 Document Scanner.
Update-MSOffice

Updates Microsoft Office desktop applications (PowerPoint, Word, Excel).
Update-Browsers

Updates Google Chrome and Mozilla Firefox browsers.
Check-OktaPlugin

Checks the installation and version of the Okta plugin in Chrome and Firefox.
Main Script Execution

The main script first checks the connectivity to the OWL device and then proceeds to call the update functions.
Digital Signing Placeholder

A placeholder is provided for digital signing using a code signing certificate.
Script

powershell

<#
.SYNOPSIS
Unified PowerShell script to update various software and manage OWL firmware.

.DESCRIPTION
This script updates MS Paint, Photos, Snip Tool, Zoom Client, Aver U50 Document Scanner app,
MS Office Desktop Apps, Chrome, Firefox, and checks the Okta Plugin. It also includes functionality to check
the network connectivity of OWL devices and placeholders for digital signing.

#>

# Parameters
param(
    [string]$OwlDeviceIP = "192.168.1.100"
)

# Function to test device connectivity
function Test-DeviceConnectivity {
    param(
        [string]$DeviceIP
    )
    $ping = Test-Connection -ComputerName $DeviceIP -Count 1 -Quiet
    return $ping
}

# Function to update applications using winget
function Update-Applications {
    # Ensure winget is installed
    if (-Not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Host "winget is not installed. Please install it first."
        exit
    }

    # Update MS Paint
    winget upgrade Microsoft.Paint

    # Update Photos
    winget upgrade Microsoft.Photos

    # Update Snip Tool
    winget upgrade Microsoft.ScreenSketch

    # Update Zoom
    winget upgrade Zoom.Zoom

    Write-Host "Applications have been updated successfully."
}

# Function to update Aver U50 Document Scanner app
function Update-AverU50 {
    $downloadUrl = "https://www.aver.com/download-link"  # Replace with actual URL
    $installerPath = "C:\Temp\AverU50Installer.exe"

    Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath
    Start-Process -FilePath $installerPath -ArgumentList "/silent" -Wait

    Write-Host "Aver U50 Document Scanner app has been updated."
}

# Function to update MS Office Desktop Apps
function Update-MSOffice {
    Start-Process -FilePath "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" -ArgumentList "updatestore=true" -Wait
    Write-Host "MS Office apps have been updated."
}

# Function to update browsers
function Update-Browsers {
    # Update Chrome
    winget upgrade Google.Chrome

    # Update Firefox
    winget upgrade Mozilla.Firefox

    Write-Host "Browsers have been updated successfully."
}

# Function to check Okta Plugin
function Check-OktaPlugin {
    # For Chrome
    $chromeExtensionId = "obkldkecplahdcfoenkbldakmgfoajdm"  # Example extension ID for Okta
    Start-Process "chrome://extensions/?id=$chromeExtensionId"

    # For Firefox
    $firefoxExtensionUrl = "https://addons.mozilla.org/en-US/firefox/addon/okta-secure-web-authentication/"
    Start-Process "firefox.exe" -ArgumentList $firefoxExtensionUrl

    Write-Host "Check the browser extension pages to ensure the Okta plugin is up to date."
}

# Main script execution
if (Test-DeviceConnectivity -DeviceIP $OwlDeviceIP) {
    Write-Host "OWL Device is reachable. Proceed with the firmware update using the OWL app."
} else {
    Write-Host "OWL Device is not reachable. Check the network connection and try again."
}

# Call update functions
Update-Applications
Update-AverU50
Update-MSOffice
Update-Browsers
Check-OktaPlugin

# Placeholder for digital signing
# To sign the script, you will need a code signing certificate.
# Once you have the certificate, you can use the Set-AuthenticodeSignature cmdlet to sign the script.
# Example:
# $cert = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert | Select-Object -First 1
# Set-AuthenticodeSignature -FilePath "path-to-your-script.ps1" -Certificate $cert

Write-Host "All updates and checks are complete."

Digital Signing

To ensure the script is secure and trusted, follow these steps to sign the script using a code signing certificate.

    Obtain a Code Signing Certificate:
        Obtain a certificate from a trusted certificate authority (CA) or generate a self-signed certificate for internal use.

    Sign the Script:
        Use the Set-AuthenticodeSignature cmdlet to sign the script:

    powershell

    $cert = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert | Select-Object -First 1
    Set-AuthenticodeSignature -FilePath "path-to-your-script.ps1" -Certificate $cert

Converting to Executable

To convert the PowerShell script into an executable (.exe) file, use the following steps:

    Install PS2EXE:
        Install the ps2exe module from the PowerShell Gallery:

    powershell

Install-Module -Name PS2EXE -Scope CurrentUser

Convert Script to Executable:

    Use the ps2exe tool to convert the PowerShell script to an executable file:

powershell

    ps2exe -inputFile "path-to-your-script.ps1" -outputFile "path-to-your-script.exe"

Usage

    Run the Script:
        Execute the PowerShell script directly:

    powershell

.\your-script.ps1

Run the Executable:

    Once converted to an executable, run it as any other application:

cmd

    your-script.exe

Notes

    Always ensure you have the correct administrative privileges to perform these updates.
    Test the script in a controlled environment before deploying it widely.
    Maintain a backup of critical systems and data before running updates.
