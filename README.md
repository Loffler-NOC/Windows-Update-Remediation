# Windows-Update-Remediation
Attempts to fix failing windows updates

Steps this takes:
- Running sfc scannow
- Making sure wuaserv (Windows Update service) is set to start up automatically
- Stopping Windows Update service and components
- Deleting qmgr.dat files
- Registering BITS files
- Resetting winsock
- Resetting catroot2 folder
- Starting Windows Update service and components
- Attempting to install Windows updates
