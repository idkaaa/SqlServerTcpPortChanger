<# 
.Synopsis 
  Load-Smo-Assembly tries to load the correct SMO assembly for a particular system's SQL Server version.
.DESCRIPTION 
  This function will attempt to load the specified SMO dlls from a provided directory. It
  will return a valid WMI object if it finds an instance or $null if it doesn't.
.NOTES 
   Created by: Christopher Hiles
   Modified: 2017/02/07 

.PARAMETER SmoDirectory 
   This is the directory that contains the SMO dlls without trailing slash.
.EXAMPLE 
   Load-Smo-Assembly -SmoDirectory "C:\Tools\SMO"  
   Attempts to load specified dlls from C:\Tools\SMO
#>
$ScriptDirectory = Split-Path -parent $PSCommandPath
# Load Dependencies
. "$ScriptDirectory\Logger.ps1"

function Load-Smo-Assembly
{
	[CmdletBinding()]
	[OutputType([System.Management.Automation.PSObject])]
    Param 
    ( 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [string]$SmoDirectory
	);

    Begin
    {
        $SmoAssemblyPathIndex = 0;
		$SmoAssemblyPathList = @(
			"$SmoDirectory\Microsoft.SqlServer2016.SqlWmiManagement.dll",
			"$SmoDirectory\Microsoft.SqlServer2014.SqlWmiManagement.dll",
			"$SmoDirectory\Microsoft.SqlServer2012.SqlWmiManagement.dll",
			"$SmoDirectory\Microsoft.SqlServer2008R2.SqlWmiManagement.dll"
			);

		#2017-02-07 CLH
		# Function to load a specific assembly dll.
		# http://stackoverflow.com/a/37468429
		function f_LoadSmoFromAssembly ([string]$AssemblyPath)
		{
			$SmoAssemblyName = "Microsoft.SqlServer.SqlWmiManagement";
			$SmoAssemblyTypeName = "Microsoft.SqlServer.Management.Smo.Wmi.ManagedComputer";
			Write-Log "Loading $SmoAssemblyName from $AssemblyPath";
			$bytes = [System.IO.File]::ReadAllBytes($AssemblyPath);
			$SmoAssembly = [System.Reflection.Assembly]::Load($bytes);
			$SpecificTypeName = $SmoAssembly.GetType($SmoAssemblyTypeName).AssemblyQualifiedName;
			return New-Object ($SpecificTypeName);
		}
    };

    Process
    {
        [bool]$FoundSqlInstance = $false;
		while ($FoundSqlInstance -eq $false)
		{
			if($SmoAssemblyPathIndex -eq $SmoAssemblyPathList.Count)
			{
				Write-Log "Error: No valid assemblies found in list.";
				return;
			}
			$SmoAssemblyPath = $SmoAssemblyPathList[$SmoAssemblyPathIndex++];
			Write-Log "Attempting to load SMO assembly: $SmoAssemblyPath";
			$SqlManagementObject = f_LoadSmoFromAssembly -AssemblyPath $SmoAssemblyPath;
			if($SqlManagementObject.ServerInstances.Count -gt 0)
			{
				$FoundSqlInstance = $true;
			}
			else
			{
				Write-Log "Failed to find SQL Instances using Assembly: $SmoAssemblyPath";
			}
		}
    };

    End
    {
		if($FoundSqlInstance -eq $true)
		{
			Write-Log "Found an instance. Returning loaded Smo object using: $SmoAssemblyPath";
			return $SqlManagementObject;
		}
		else
		{
			Write-Log "Assemblies failed to find any instances.";
			return $null;
		}
    };
};
