# 2017/02/07 CLH
# Main.ps1
#

$ScriptDirectory = Split-Path -parent $PSCommandPath;
# Load Dependencies
. "$ScriptDirectory\Logger.ps1";
. "$ScriptDirectory\SetSqlServerIpAllTcpPort.ps1";
Add-Type -AssemblyNamePresentationFramework

Write-Log "Starting Sql Setup";
Write-Log "Setting SQL Tcp Port to 1433";
[bool] $SuccessTcpSet = Set-SqlServerIpAllTcpPort -IpAllTcpPort "1433" -InstanceName "SQLEXPRESS";
if($SuccessTcpSet -eq $false)
{
	[System.Windows.MessageBox]::Show("Error: Failed to set TCP port. Check log in: $ScriptDirectory");
}