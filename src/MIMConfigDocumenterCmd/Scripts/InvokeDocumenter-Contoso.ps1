<#
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.

	This is a sample script demonstrating use of PowerShell command-line interface of MIM Configuration Documenter.

	Pre-requisites:
		.NET Framework 4.5 is required to run this tool.
		In addition, to generate report for FIMService config, FIMAutomation PowerShell snap-in must be installed on the local server.
		The FIMAutomation PowerShell snap-in is installed by the the MIM/FIM installer when the FIMService role is selected.
		If you are running this tool on your desktop, you must install the FIMAutomation PowerShell snap-in manually
		by following in instructions provided in the Documenter Wiki.

	To use this script: 
		Place the production FIM / MIM Service schema and policy exports file in the Data\<Customer>\<Production>\ServiceConfig directory.
		Place the production FIM / MIM Sync server configuration exports file in the Data\<Customer>\<Production>\SyncConfig directory.
		Place the pilot FIM / MIM Service schema and policy exports file in the Data\<Customer>\<Pilot>\ServiceConfig directory.
		Place the pilot FIM / MIM Sync server configuration exports file in the Data\<Customer>\<Pilot>\SyncConfig directory.
		Edit the $pilotConfig and $productionConfig (and $reportType) variables appropriately.
#>

Set-StrictMode -Version "2.0"

######## Edit as appropriate ####################################
$pilotConfig = "Contoso\Pilot" # the path of the Pilot / Target config export files relative to the MIM Configuration Documenter "Data" folder.
$productionConfig = "FIM-R2-SP1-Base_4.1.3419.0" # the path of the Production / Baseline config export files relative to the MIM Configuration Documenter "Data" folder.
$reportType = "SyncAndService" # "SyncOnly" # "ServiceOnly"
#################################################################

$global:VerbosePreference = "Continue"
$documenterModuleName = "MIMConfigDocumenter"

$hostSettings = (Get-Host).PrivateData
$hostSettings.WarningBackgroundColor = "red"
$hostSettings.WarningForegroundColor = "white"

Set-Location (Split-Path -Path $MyInvocation.MyCommand.Definition -Parent)

if (Get-Module -Name $documenterModuleName)
{
    Remove-Module -Name $documenterModuleName
}

Import-Module -Name (Join-Path -Path $PWD -ChildPath "$documenterModuleName.psm1") -ErrorAction Stop

Write-Host "Invoking cmdlet: Get-MIMConfigReport..."
Get-MIMConfigReport -PilotConfig $pilotConfig -ProductionConfig $productionConfig -ReportType $reportType
Write-Host "Invokation complete! Please check any errors or warnings in the MIMConfigDocumenter-Error.log..."

Read-Host "Press any key to exit"
