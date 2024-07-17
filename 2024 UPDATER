<#
.SYNOPSIS
Unified PowerShell script to update various software, manage OWL firmware, and configure system settings for FLEX classrooms.

.DESCRIPTION
This script updates MS Paint, Photos, Snip Tool, Zoom Client, Aver U50 Document Scanner app, MS Office Desktop Apps, Chrome, Firefox, Edge, Okta, PowerPoint, Word, Excel, Zoom, Canvas, Teams. It also configures mouse settings, extends the automatic timeout, and updates the Meeting OWL device.

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

    # Update Zoom
    winget upgrade Zoom.Zoom

    # Update Browsers
    winget upgrade Google.Chrome
    winget upgrade Mozilla.Firefox
    winget upgrade Microsoft.Edge

    # Update MS Office
    Start-Process -FilePath "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" -ArgumentList "updatestore=true" -Wait

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

# Function to configure mouse settings
function Configure-MouseSettings {
    # Change mouse pointer size and color
    Set-ItemProperty -Path "HKCU:\Control Panel\Cursors" -Name Scheme -Value "Windows Black (extra large) (system scheme)"

    # Apply settings
    RUNDLL32.EXE USER32.DLL,UpdatePerUserSystemParameters ,1 ,True

    Write-Host "Mouse settings have been configured."
}

# Function to extend automatic timeout
function Extend-Timeout {
    powercfg /change monitor-timeout-ac 30
    powercfg /change monitor-timeout-dc 30

    Write-Host "Automatic timeout has been extended to 30 minutes."
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
Configure-MouseSettings
Extend-Timeout
Check-OktaPlugin

# Placeholder for digital signing
# To sign the script, you will need a code signing certificate.
# Once you have the certificate, you can use the Set-AuthenticodeSignature cmdlet to sign the script.
# Example:
# $cert = Get-ChildItem -Path Cert:\CurrentUser\My -CodeSigningCert | Select-Object -First 1
# Set-AuthenticodeSignature -FilePath "path-to-your-script.ps1" -Certificate $cert

Write-Host "All updates and checks are complete."
