# MIM Configuration Documenter

MIM configuration documenter is a tool to generate documentation of a MIM / FIM synchronization or service installation.

The goal of this project is to:

* Document deployment configuration details for the MIM / FIM solution!
* Track any configuration changes you have made since a specific baseline!!
* Build confidence in getting things right when making changes to the deployed solution!!

Prerequisites:

1. .NET Framework 4.5 to be able to run the tool.
2. FIMAutomation PowerShell snap-in installed locally when generating report for FIMService config.
3. A modern browser (e.g. Microsoft Edge) to view the report.

How to use the tool:

* Download the latest release MIMConfigDocumenter.zip from the [releases](https://github.com/Microsoft/MIMConfigDocumenter/releases) tab under the Code tab tab, UNBLOCK the downloaded zip file and extract the zip file to an empty local folder.
	* This will extract the Documenter application binaries along with the sample data files for "Contoso".
	* Make sure that the tool runs by running the PowerShell script InvokeDocumenter-Contoso.ps1.
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
