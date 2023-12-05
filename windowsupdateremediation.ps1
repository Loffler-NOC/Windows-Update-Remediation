<#Check if ExecutionPolicy is set to Unrestricted
$ExecutionPolicy = Get-ExecutionPolicy

if ($ExecutionPolicy -eq "Unrestricted") {
    Write-Host "Execution policy is set to Unrestricted"
} else {
    Write-Host "Execution policy is not set to Unrestricted. Attempting to set it to Unrestricted now."
    Set-ExecutionPolicy Unrestricted
}
#>

#Run an sfc /scannow in case anything is corrupted
sfc /scannow

#Make sure that the startup type of the windows update service is automatic
$serviceName = "wuauserv"
$service = Get-Service -Name $serviceName

if ($service -ne $null) {
    if ($service.StartType -ne "Automatic") {
        Set-Service -Name $serviceName -StartupType Automatic
        Write-Host "The startup type of service '$serviceName' has been set to Automatic."
    }
    else {
        Write-Host "The startup type of service '$serviceName' is already set to Automatic."
    }
}
else {
    Write-Host "The service '$serviceName' was not found."
}


# Stop Windows Update service
Write-Host "Stopping Windows Update Service"
Stop-Service -Name "wuauserv" -Force

# Stop the Windows Update components
$script = {
    $ErrorActionPreference = "Stop"
    $services = "CryptSvc", "BITS", "msiserver", "wuauserv"
    foreach ($service in $services) {
        Write-Host "Stopping service: " $service
        Stop-Service -Name $service -Force
    }
}
Invoke-Command -ScriptBlock $script

#Delete qmgr*.dat files
Write-Host "Removing qmgr*.dat files"
Remove-Item -Path "${env:ALLUSERSPROFILE}\Application Data\Microsoft\Network\Downloader\qmgr*.dat"

#Set location to system 32 folder
Write-Host "Setting location to ${env:windir}\system32"
Set-Location -Path "${env:windir}\system32"


#Regester BITS files
Write-Host "Registering BITS files"
$files = @(
    "atl.dll",
    "urlmon.dll",
    "mshtml.dll",
    "shdocvw.dll",
    "browseui.dll",
    "jscript.dll",
    "vbscript.dll",
    "scrrun.dll",
    "msxml.dll",
    "msxml3.dll",
    "msxml6.dll",
    "actxprxy.dll",
    "softpub.dll",
    "wintrust.dll",
    "dssenh.dll",
    "rsaenh.dll",
    "gpkcsp.dll",
    "sccbase.dll",
    "slbcsp.dll",
    "cryptdlg.dll",
    "oleaut32.dll",
    "ole32.dll",
    "shell32.dll",
    "initpki.dll",
    "wuapi.dll",
    "wuaueng.dll",
    "wuaueng1.dll",
    "wucltui.dll",
    "wups.dll",
    "wups2.dll",
    "wuweb.dll",
    "qmgr.dll",
    "qmgrprxy.dll",
    "wucltux.dll",
    "muweb.dll",
    "wuwebv.dll"
)

foreach ($file in $files) {
    $regsvr32Path = Join-Path -Path $env:SystemRoot -ChildPath "System32\regsvr32.exe"
    Write-Host "regsvr32Path is $regsvr32Path"
    $filePath = Join-Path -Path $env:SystemRoot -ChildPath "System32\$file"
    Write-Host "filePath is $filePath"
    Write-Host "Registering $file"
    Start-Process -FilePath $regsvr32Path -ArgumentList "/s", $filePath -Wait
    Write-Host "Done registering $file"
}


#Reset winsock
Write-Host "Resetting winsock"
netsh winsock reset

$random = Get-Random -Minimum 1 -Maximum 1000001


# Reset the catroot2 folder
Rename-Item -Path "$env:Windir\System32\catroot2" -NewName "catroot$random.old"

# Start Windows Update service
Write-Host "Starting Windows Update Service"
Start-Service -Name "wuauserv"


# Start the Windows Update components
$script = {
    $ErrorActionPreference = "Stop"
    $services = "CryptSvc", "BITS", "msiserver", "wuauserv"
    foreach ($service in $services) {
        Write-Host "Starting service: " $service
        Start-Service -Name $service
    }
}
Invoke-Command -ScriptBlock $script


#Attempt to install Windows Updates

# Check if NuGet package provider is installed
$providerName = "NuGet"
$providerInstalled = Get-PackageProvider -Name $providerName -ErrorAction SilentlyContinue -ForceBootstrap

if ($providerInstalled) {
    Write-Host "NuGet package provider is installed."
} else {
    Write-Host "NuGet package provider is not installed. Attemping to install it now."
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
}

# Check if the PSWindowsUpdate module is already installed
$moduleName = "PSWindowsUpdate"
$installedModules = Get-Module -ListAvailable | Where-Object { $_.Name -eq $moduleName }

if ($installedModules) {
    Write-Host "The '$moduleName' module is installed."
} else {
    Write-Host "The '$moduleName' module is not installed. Attemping to install it now."
    Install-Module -Name PSWindowsUpdate -Force
}

Write-Host "Importing PSWindowsUpdate module"
Import-Module PSWindowsUpdate
Write-Host "Attempting to install Windows Updates"
Install-WindowsUpdate -AcceptAll -IgnoreReboot
Write-Host "Done attempting to install Windows Updates"
