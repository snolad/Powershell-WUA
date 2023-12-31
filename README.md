# Powershell-WUA

Use this Powershell module to exploit Windows Update Agent Win32 API on your Windows system (WinX/WinServerX). This module contains functions that give you the abiilty to do following operations :
* Search online missing (not installed) updates on the computer using a specified online service with the Windows Update Agent, and by update category.
* Search missing (not installed) updates on the computer using a registered offline scan service (using offline catalog cab file) with the Windows Update Agent.
* Install specified list of updates set in paramater.
* Remove an update service by his ID.
* Register a Microsoft Update Service for offline scan using your offline catalog cab file.
* Download from Intenet the last up-to-date version of the Microsoft offline catalog, required for offline scan.

## Installation

To install this module, use the following command :

```powershell
Install-Module ".\PSWindowsUpdateAgent.psm1"
```

## Search-MissingUpdatesOnline()
```powershell
Search-MissingUpdatesOnline -UpdateService MicrosoftUpdate -UpdateCategory CriticalUpdates
```

## Search-MissingUpdatesOffline()
```powershell
Search-MissingUpdatesOffline -ServiceId 'ServiceId'
```

## Install-Updates()
```powershell
Install-Updates -MissingUpdatesCollection $yourMissingUpdates
```

## Register-OfflineUpdateService()
```powershell
Register-OfflineUpdateService -OfflineCatalogPath "YourPath\wsusscn2.cab"
```

## Get-OfflineCatalogFromInternet()
```powershell
Get-OfflineCatalogFromInternet -DestinationFolderPath "YourPath"
```

## Remove-UpdateService()
```powershell
Remove-UpdateService -ServiceId 'ServiceId'
```

## License & Contributing
The PSWindowsUpdateAgent module is licensed under MIT. Any contributions highly appreciated!
