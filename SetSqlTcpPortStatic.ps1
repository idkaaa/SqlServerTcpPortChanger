# 2017-02-07 CLH
# Set all TCP ports for Sql Server to listen on port 1433.
#

$ScriptDirectory = Split-Path -parent $PSCommandPath

# Load logging script
. "$ScriptDirectory\Logger.ps1"

$SqlInstanceName = "SQLEXPRESS"

$AlternateSmoAssemblyPathIndex = 0
$AlternateSmoAssemblyPathList = @(
	"$ScriptDirectory\SMO\Microsoft.SqlServer2014.SqlWmiManagement.dll",
	"$ScriptDirectory\SMO\Microsoft.SqlServer2016.SqlWmiManagement.dll",
	"$ScriptDirectory\SMO\Microsoft.SqlServer2012.SqlWmiManagement.dll",
	"$ScriptDirectory\SMO\Microsoft.SqlServer2008R2.SqlWmiManagement.dll"
	)

#2017-02-07 CLH
# Function to load a specific assembly dll.
# http://stackoverflow.com/a/37468429
function p_LoadSmoFromAssembly ([string]$AssemblyPath)
{
	$SmoAssemblyName = "Microsoft.SqlServer.SqlWmiManagement"
	$SmoAssemblyTypeName = "Microsoft.SqlServer.Management.Smo.Wmi.ManagedComputer"
	Write-Log "Loading $SmoAssemblyName from $AssemblyPath"
	$bytes = [System.IO.File]::ReadAllBytes($AssemblyPath)
	$SmoAssembly = [System.Reflection.Assembly]::Load($bytes)
	$SpecificTypeName = $SmoAssembly.GetType($SmoAssemblyTypeName).AssemblyQualifiedName
	return New-Object ($SpecificTypeName)
}

[bool]$ConnectedToSqlInstance = $false
while ($ConnectedToSqlInstance -eq $false)
{
	$SmoAssemblyPath = $AlternateSmoAssemblyPathList[$AlternateSmoAssemblyPathIndex++]
	Write-Log "Attempting to load SMO assembly: $SmoAssemblyPath"
	$SqlManagementObject = p_LoadSmoFromAssembly -AssemblyPath $SmoAssemblyPath
	try
	{
		Write-Log "Attempting to access SQL Instance Managed Computer Object."
		$SqlServerTcpSettings = $SqlManagementObject.ServerInstances[$SqlInstanceName].ServerProtocols['Tcp']
		$TcpPortStaticProperties = $SqlServerTcpSettings.IPAddresses['IPAll'].IPAddressProperties['TcpPort']
		Write-Log "Successfully connected to sql instance."
		$ConnectedToSqlInstance = $true
		Write-Log "Setting IPAll TCP port to 1433."
		$TcpPortStaticProperties.Value = '1433'
		Write-Log "Restarting sql server."
		$SqlServerTcpSettings.Alter()
	}
	catch [System.Management.Automation.RuntimeException]
	{
		Write-Log "Failed to connect to the SQL instance: $SqlInstanceName"
		Write-Log "<!<System.Management.Automation.RuntimeException>!>"
		Write-Log $_.Exception.Message
		Write-Log "Is Sql Server Installed on $Env:COMPUTERNAME\$SqlInstanceName ?"
		if($AlternateSmoAssemblyPathList.Count -eq $AlternateSmoAssemblyPathIndex)
		{
			Write-Log "Error: Failed to load any valid SMO assembly. Exiting.";
			exit
		}
	}
}


