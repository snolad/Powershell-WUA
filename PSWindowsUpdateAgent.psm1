<#PSScriptInfo

.NOTES 
	Version: 1.0.0
	Date : 2023-10-01
	Author: Steeve DADON
	Licence: MIT

.DESCRIPTION
Powershell module for searching and installing Windows Updates

.GUID 9b0c3246-6ad9-496c-98ca-861201579fe5

.AUTHOR Steeve DADON

.COPYRIGHT (c) 2023 Steeve DADON

.LICENSEURI https://github.com/snolad/Powershell-WUA/blob/main/LICENCE

.PROJECTURI https://github.com/snolad/Powershell-WUA/

.RELEASENOTES

#>

function Get-OfflineCatalogFromInternet
{
	<#
		.SYNOPSIS
		Download from Intenet the last up-to-date version of the Microsoft offline catalog, required for offline scan.

		.PARAMETER -DestinationFolderPath
		System.String
		Specify the destination folder path for the downloaded file.

		.EXAMPLE
		Get-OfflineCatalogFromInternet -DestinationFolderPath "c:\temp"

		.OUTPUTS
		System.String
		This function outputs a string containing the offline catalog file path downloaded.
	#>
	
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)][string]$DestinationFolderPath
	)
	
	begin
	{
		$ErrorActionPreference = 'Stop'
		Write-Host "Calling function $($MyInvocation.MyCommand.Name)"
		$WINDOWSOFFLINECATALOG_URL = "https://go.microsoft.com/fwlink/?LinkID=74689"
	}

	process
	{
		try
		{
			$DestinationFolderPath = $DestinationFolderPath.TrimEnd('\')
			if ((Test-Path -Path $DestinationFolderPath -PathType Container) -eq $false)
			{
				Throw "Destination Path is not a valid directory or doesn't exit"
			}
			else
			{
				$OutputFilePath = "{0}\{1}" -f $DestinationFolderPath, "wsusscn2.cab"
			}

			Write-Host "Downloading of $WINDOWSOFFLINECATALOG_URL using .Net Web Client..."
			$webclient = New-Object -TypeName System.Net.WebClient
			$webclient.UseDefaultCredentials = $true
			$webclient.DownloadFile($WINDOWSOFFLINECATALOG_URL, $OutputFilePath)
			Write-Host "Download of $OutputFilePath in success" -ForegroundColor Green
		}
		catch
		{
			Write-Error $_.Exception.Message
		}
	}

	end
	{
		$webclient.Dispose()
		return $OutputFilePath
	}
}

function Register-OfflineUpdateService
{
	<#
		.SYNOPSIS
		Register a Microsoft Update Service for offline scan using your offline catalog cab file.

		.PARAMETER -OfflineCatalogPath
		System.String
		Specify the location of your offline catalog cab file.

		.EXAMPLE
		Register-OfflineUpdateService -OfflineCatalogPath "c:\temp\wsusscn2.cab"

		.OUTPUTS
		System.__ComObject
		This function outputs a ComObject of type IUpdateService.

		.NOTES
		This service registration is needed in order to run an offline scan using Search-MissingUpdatesOffline.
		To run an Offline scan using Search-MissingUpdatesOffline function, you will need to specify the ServiceID property value returned by the object of this function.
		Administrator permission is needed to run this function.
	#>
	
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$OfflineCatalogPath
	)

	begin
	{
		$ErrorActionPreference = 'Stop'
		Write-Host "Calling function $($MyInvocation.MyCommand.Name)"

		#Administrator permission check
		if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
		{
			Write-Error "Not running as administrator. Cannot continue process..."
		}
	}

	process
	{
		try
		{
			if ((Test-Path -Path $OfflineCatalogPath) -eq $false)
			{
				Throw "$OfflineCatalogPath is not valid or doesn't exists"
			}
	
			# Create Offline Update Service
			$objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
			$offlineUpdateService = $objServiceManager.AddScanPackageService("Offline Update Service", $OfflineCatalogPath, 1)
			Write-Host "Offline update service registered in success" -ForegroundColor Green
		}
		catch
		{
			Write-Error $_.Exception.Message
		}
	}
	end
	{
		return $offlineUpdateService
	}
}

function Remove-UpdateService
{
	<#
		.SYNOPSIS
		Remove an update service by his ID.

		.PARAMETER -ServiceId
		System.String
		Specify the service ID to remove.

		.EXAMPLE
		Remove-UpdateService -ServiceId '3d39f8ff-fb03-4bfb-84fe-97146ecf789e'

		.NOTES
		Administrator permission is needed to run this function.
	#>
	
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$ServiceId
	)

	begin
	{
		$ErrorActionPreference = 'Stop'
		Write-Host "Calling function $($MyInvocation.MyCommand.Name)"

		#Administrator permission check
		if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
		{
			Write-Error "Not running as administrator. Cannot continue process..."
		}
	}
	
	process
	{
		try
		{
			$objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager"
			$objServiceManager.RemoveService($ServiceId)
			Write-Host "Update service with ID $ServiceId successfully removed" -ForegroundColor Green
		}
		catch
		{
			Write-Error $_.Exception.Message
		}
	}
}

function Search-MissingUpdatesOnline
{
	<#
		.SYNOPSIS
		Search missing (not installed) updates on the computer using a specified online service with the Windows Update Agent, and by update category.

		.PARAMETER -UpdateService
		System.String
		Specify the an update service according to following values "Default", "MicrosoftUpdate", "WindowsUpdate".
		
		.PARAMETER -UpdateCategory
		Specify a category of update accoding to the list of choices, by default the search is made with all categories.

		.EXAMPLE
		Search-MissingUpdatesOnline -UpdateService MicrosoftUpdate -UpdateCategory CriticalUpdates

		.OUTPUTS
		Object[]
		This function outputs is an array of Object of type ISearchResult, which is a list of Updates object.

		.NOTES
		The output of this function can be used as parameter of the function Install-Updates.
	#>
	
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[ValidateSet("Default", "MicrosoftUpdate", "WindowsUpdate")]
		$UpdateService,
		[Parameter(Mandatory = $false)]
		[ValidateSet("Application", "Connectors", "CriticalUpdates", "DefinitionUpdates", "DeveloperKits", "FeaturePacks", "Guidance", "SecurityUpdates", "ServicePacks", "Tools", "UpdateRollups", "Updates")]
		$UpdateCategory
	)

	begin
	{
		$ErrorActionPreference = 'Stop'
		Write-Host "Calling function $($MyInvocation.MyCommand.Name)"

		#Update service GUID constants from Microsoft documentation
		$WINDOWSUPDATESERVICE_GUID = '9482f4b4-e343-43b6-b170-9a65bc822c77'
		$MICROSOFTUPDATESERVICE_GUID = '7971f918-a847-4430-9279-4a52d1efe18d'
		
		$criteria = "IsInstalled=0"

		if ($UpdateCategory)
		{
			switch ($UpdateCategory)
			{
				# Sample list of Update category IDs
				Application { $updateCategoryID = '5C9376AB-8CE6-464A-B136-22113DD69801' }
				Connectors { $updateCategoryID = '434DE588-ED14-48F5-8EED-A15E09A991F6' }
				CriticalUpdates { $updateCategoryID = 'E6CF1350-C01B-414D-A61F-263D14D133B4' }
				DefinitionUpdates { $updateCategoryID = 'E0789628-CE08-4437-BE74-2495B842F43B' }
				DeveloperKits { $updateCategoryID = 'E140075D-8433-45C3-AD87-E72345B36078' }
				FeaturePacks { $updateCategoryID = 'B54E7D24-7ADD-428F-8B75-90A396FA584F' }
				Guidance { $updateCategoryID = '9511D615-35B2-47BB-927F-F73D8E9260BB' }
				SecurityUpdates { $updateCategoryID = '0FA1201D-4330-4FA8-8AE9-B877473B6441' }
				ServicePacks { $updateCategoryID = '68C5B0A3-D1A6-4553-AE49-01D3A7827828' }
				Tools { $updateCategoryID = 'B4832BD8-E735-4761-8DAF-37F882276DAB' }
				UpdateRollups { $updateCategoryID = '28BC880E-0592-4CBF-8F95-C79B17911D5F' }
				Updates { $updateCategoryID = 'CD5FFD1E-E932-4E3A-BF74-18BF0B1BBD8' }
			}
			$criteria += " and CategoryIDs contains '$updateCategoryID'"
		}

		#Object Searcher intialization
		$objSession = New-Object -ComObject "Microsoft.Update.Session"
		$objSearcher = $objSession.CreateUpdateSearcher()
	}

	process
	{
		try 
		{
			if ($UpdateService -eq "Default")
			{
				$objSearcher.ServerSelection = 0
			}
			elseif ($UpdateService -eq "WindowsUpdate")
			{
				$objSearcher.ServerSelection = 2
				$objSearcher.ServiceID = $WINDOWSUPDATESERVICE_GUID
			}
			elseif ($UpdateService -eq "MicrosoftUpdate")
			{
				$objSearcher.ServerSelection = 3
				$objSearcher.ServiceID = $MICROSOFTUPDATESERVICE_GUID
			}
			
			Write-host "Searching for missing updates using update service [$UpdateService]..." -NoNewline
			$stopwatch = [system.diagnostics.stopwatch]::StartNew()
			$updates = $objSearcher.Search($criteria)
			$stopwatch.Stop()
			$elapsed = "{0:hh\:mm\:ss}" -f $stopwatch.Elapsed
			Write-Host " OK - Scan duration : $elapsed " -ForegroundColor Green
		}
		catch 
		{
			Write-Error $($_.Exception.Message)
		}
	}
	end
	{
		Write-host "Found $($updates.Updates.Count) missing update(s)"
		return $updates.Updates
	}
}

function Search-MissingUpdatesOffline
{
	<#
		.SYNOPSIS
		Search missing (not installed) updates on the computer using a registered offline scan service with the Windows Update Agent.

		.PARAMETER -ServiceId
		System.String
		Specify the offline scan service ID to use.
		
		.PARAMETER -UpdateCategory
		Specify a category of update accoding to the list of choices, by default the search is made with all categories.

		.EXAMPLE
		Search-MissingUpdatesOffline -ServiceId 'xxxxxxxxxx'

		.OUTPUTS
		Object[]
		This function outputs is an array of Object of type ISearchResult, which is a list of Updates object.

		.NOTES
		The output of this function can be used as parameter of the function Install-Updates.
	#>

	
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[string]$ServiceId,
		[Parameter(Mandatory = $false)]
		[ValidateSet("Application", "Connectors", "CriticalUpdates", "DefinitionUpdates", "DeveloperKits", "FeaturePacks", "Guidance", "SecurityUpdates", "ServicePacks", "Tools", "UpdateRollups", "Updates")]
		$UpdateCategory
	)

	begin
	{
		$ErrorActionPreference = 'Stop'
		Write-Host "Calling function $($MyInvocation.MyCommand.Name)"
		
		$criteria = "IsInstalled=0"

		if ($UpdateCategory)
		{
			switch ($UpdateCategory)
			{
				# Sample list of Update category IDs
				Application { $updateCategoryID = '5C9376AB-8CE6-464A-B136-22113DD69801' }
				Connectors { $updateCategoryID = '434DE588-ED14-48F5-8EED-A15E09A991F6' }
				CriticalUpdates { $updateCategoryID = 'E6CF1350-C01B-414D-A61F-263D14D133B4' }
				DefinitionUpdates { $updateCategoryID = 'E0789628-CE08-4437-BE74-2495B842F43B' }
				DeveloperKits { $updateCategoryID = 'E140075D-8433-45C3-AD87-E72345B36078' }
				FeaturePacks { $updateCategoryID = 'B54E7D24-7ADD-428F-8B75-90A396FA584F' }
				Guidance { $updateCategoryID = '9511D615-35B2-47BB-927F-F73D8E9260BB' }
				SecurityUpdates { $updateCategoryID = '0FA1201D-4330-4FA8-8AE9-B877473B6441' }
				ServicePacks { $updateCategoryID = '68C5B0A3-D1A6-4553-AE49-01D3A7827828' }
				Tools { $updateCategoryID = 'B4832BD8-E735-4761-8DAF-37F882276DAB' }
				UpdateRollups { $updateCategoryID = '28BC880E-0592-4CBF-8F95-C79B17911D5F' }
				Updates { $updateCategoryID = 'CD5FFD1E-E932-4E3A-BF74-18BF0B1BBD8' }
			}
			$criteria += " and CategoryIDs contains '$updateCategoryID'"
		}

		#Object Searcher intialization
		$objSession = New-Object -ComObject "Microsoft.Update.Session"
		$objSearcher = $objSession.CreateUpdateSearcher()
	}

	process
	{
		try 
		{
			$objSearcher.ServerSelection = 3
			$objSearcher.ServiceID = $ServiceId
			
			Write-host "Searching for missing updates using update service Offline scan service..." -NoNewline
			$stopwatch = [system.diagnostics.stopwatch]::StartNew()
			$updates = $objSearcher.Search($criteria)
			$stopwatch.Stop()
			$elapsed = "{0:hh\:mm\:ss}" -f $stopwatch.Elapsed
			Write-Host " OK - Scan duration : $elapsed " -ForegroundColor Green
		}
		catch 
		{
			Write-Error $($_.Exception.Message)
		}
	}
	end
	{
		Write-host "Found $($updates.Updates.Count) missing update(s)"
		return $updates.Updates
	}
}

function Install-Updates
{
	<#
		.SYNOPSIS
		Install specified list of updates set in paramater

		.PARAMETER -MissingUpdatesCollection
		System.Array
		Array of update objects of type ISearchResult, returned by function Search-MissingUpdatesOffline or by function Search-MissingUpdatesOnline

		.EXAMPLE
		Install-Updates -MissingUpdatesCollection $missingUpdates

		.OUTPUTS
		__ComObject
		Return an object of type IInstallationResult which contain the result of installations
	#>

	
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true)]
		[System.Array]$MissingUpdatesCollection
	)
	
	begin
	{
		$ErrorActionPreference = 'Stop'
		Write-Host "Calling function $($MyInvocation.MyCommand.Name)"

		if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
		{
			Write-Error "Not running as administrator. Cannot continue process..."
		}
	}

	process
	{
		try
		{
			$objSession = New-Object -ComObject "Microsoft.Update.Session"
			$objUpdateColl = New-Object -ComObject "Microsoft.Update.UpdateColl"
		
			foreach ($update in $missingUpdatesCollection)
			{
				$objUpdateColl.Add($update) | Out-Null
			}

			$downloader = $objSession.CreateUpdateDownloader()
			$downloader.Updates = $objUpdateColl
			
			Write-host "Start downloading update..."
			$downloadResult = $downloader.Download()
			Write-host "Download process finish"
			
			Write-host "Check ResultCode" -NoNewline
			switch -exact ($downloadResult.ResultCode)
			{
				0 { $status = "NotStarted" }
				1 { $status = "InProgress" }
				2 { $status = "Downloaded" }
				3 { $status = "DownloadedWithErrors" }
				4 { $status = "Failed" }
				5 { $status = "Aborted" }
			}
			Write-Host " $status"

			if ($status -ne "Downloaded")
			{
				Write-Error "Update download status not downloaded...Process aborted."
			}

			$installer = $objSession.CreateUpdateInstaller()
			
			Write-host "Start installing update..."
			$installer.Updates = $objUpdateColl
			$installationResult = $installer.Install()
			Write-host "Install process finish"
			Write-host "Reboot required state [$($installationResult.RebootRequired)]"
		}
		catch [System.Runtime.InteropServices.COMException]
		{
			Write-Error $_.Exception.Message
		}
	}
	end
	{
		return $installationResult
	}
}