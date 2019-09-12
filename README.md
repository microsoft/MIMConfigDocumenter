# MIM Configuration Documenter

The MIM Configuration Documenter is a tool used to generate documentation of a Microsoft Identity Manager (MIM) or Forefront Identity Manager (FIM) deployemt (both the Synchronization Service as well as the Service and Portal environments).

## Project goals:

* Document deployment configuration details for the MIM / FIM solution.
* Track configuration changes you have made since a specific baseline.
* Build confidence in getting things right when making changes to the deployed solution.

## Prerequisites:

1. .NET Framework 4.5 to be able to run the tool.
2. [FIMAutomation](https://docs.microsoft.com/en-us/powershell/module/fimautomation/?view=idm-ps-2016sp1) PowerShell snap-in installed locally when generating report for FIMService config.
3. A modern browser (e.g. Microsoft Edge) to view the report.

## Obtaining and testing the tool:

* Download the latest release of your choice from the [releases](https://github.com/Microsoft/MIMConfigDocumenter/releases) section under the Code tab in the project. We recommend downloading the MIMConfigDocumenter.zip file
* Be sure to *Unblock* the downloaded zip file before extracting the contents to an empty local folder. This process will extract the MIM Configuration Documenter application binaries, along with sample data files for the Contoso corporation.
* Test the tool by executing the included PowerShell script InvokeDocumenter-Contoso.ps1.

## Using the tool:

To generate documentation, the MIM Configuration Documenter compares any provided point in time configuration export, with another. The difference between the two configurations is used to generate the report. Several baseline configurations are provided for convenience. These includes:

* FIM-Base_4.0.3684.2
* FIM-R2-SP1-Base_4.1.3419.0
* FIM-R2-SP1-Base_4.1.3461.0
* FIM-R2
* MIM-SP1-Base_4.4.1302.0
* MIM-SP1-Base_4.4.1459.0

 with one of the baseline exports provided; or 

* Export the FIM Sync Server and FIM Service Configuration of your two environments: PILOT / TARGET / END-STATE environment and PRODUCTION / BASELINE / REFERENCE environment.
	* FIM Sync Server configuration is exported using File | Export Server Configuration menu in the FIM Sync Admin console.
	* FIM Service configuration is exported using the ExportSchema.ps1 and ExportPolicy.ps1 scripts.
	* If you have configuration files for only one enviorment, you can use any one set of the FIM or MIM config files provided with the tool as your PRODUCTION / BASELINE / REFERENCE environment.
* Copy the configuration export files produced in the previous step to SyncConfig and ServiceConfig folders under the "Data" directory of the Documenter tool.
	* As an example, the "Pilot" configuration files for the customer "Contoso" are provided as a sample in the "Data\Contoso\Pilot\SyncConfig" and "Data\Contoso\Pilot\ServiceConfig" folders.
	* **NOTE:** The names of the FIM Service schema and policy export files must be schema.xml and policy.xml respectively.
* Make a copy of the InvokeDocumenter-Contoso.ps1 script, name is appropriately and then open and edit the new script using the instructions provided in the script.
* Run your script to generate the documentation report of your configuration exports. Upon successful execution, the report will be placed in the "Report" folder.

# Contributing

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
