<# 
.Synopsis 
  Set-SqlServerIpAllTcpPort tries to set the local system's sql server IPAll TCP/IP port.
.DESCRIPTION 
  Given a valid port, the function will return true if the port was set and the 
  SQL server restarted.
.NOTES 
   Created by: Christopher Hiles
   Modified: 2017/02/07 

.PARAMETER IpAllTcpPort 
   This is the desired port
.PARAMETER InstanceName
	This is the specified instance rather than SQLEXPRESS.
.EXAMPLE 
   Set-SqlServerIpAllTcpPort -IpAllTcpPort "1433"  
   Attempts to set the local SQL server's TCP port to 1433 for all IP addresses 
   and uses SQLEXPRESS as the instance name.
.EXAMPLE 
	Set-SqlServerIpAllTcpPort -IpAllTcpPort "1433" -InstanceName "MySqlInstanceName" 
	Attempts to set the local SQL server's TCP port to 1433 for all IP addresses only
	on the instance: localhost\MySqlInstanceName.
#>

$ScriptDirectory = Split-Path -parent $PSCommandPath
# Load Dependencies
. "$ScriptDirectory\Logger.ps1"
. "$ScriptDirectory\LoadSmoAssembly.ps1"


function Set-SqlServerIpAllTcpPort
{
	[CmdletBinding()]
	[OutputType([bool])]
    Param 
    ( 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [string]$IpAllTcpPort,
		
		[Parameter(Mandatory=$false)]  
        [string]$InstanceName = "SQLEXPRESS" 
	);

    Begin
    {
		
    };

    Process
    {
		$SqlManagementObject = Load-Smo-Assembly -SmoDirectory "$ScriptDirectory\Smo";
		if($SqlManagementObject -eq $null)
		{
			Write-Log "Error: Reading SQL Instance failed using assemblies in $ScriptDirectory\SMO\"
			return $false;
		}
		Write-Log "Attempting to access SQL Instance Managed Computer Object to set TCP port: $IpAllTcpPort";
		$SqlServerTcpSettings = $SqlManagementObject.ServerInstances[$InstanceName].ServerProtocols['Tcp'];
		$TcpPortStaticProperties = $SqlServerTcpSettings.IPAddresses['IPAll'].IPAddressProperties['TcpPort'];
		Write-Log "Successfully connected to sql instance.";
		Write-Log "Setting IPAll TCP port to $IpAllTcpPort";
		$TcpPortStaticProperties.Value = $IpAllTcpPort;
		Write-Log "Restarting sql server.";
		$SqlServerTcpSettings.Alter();
		if($TcpPortStaticProperties.Value -eq $IpAllTcpPort)
		{
			Write-Log "Successfully set IP All TCP port to: $IpAllTcpPort on localhost\$InstanceName";
			return $true;
		}
		else
		{
			Write-Log "Error: Failed to set IP All TCP port to: $IpAllTcpPort on localhost\$InstanceName";
			return $false;
		}
    };

    End
    {
		
    };
};
