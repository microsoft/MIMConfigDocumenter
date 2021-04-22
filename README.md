# MIM Configuration Documenter

The MIM Configuration Documenter is a tool used to generate documentation of a Microsoft Identity Manager (MIM) or Forefront Identity Manager (FIM) deployment (both the Synchronization Service as well as the Service and Portal environments).

## Project goals:

* Document deployment configuration details for the MIM / FIM solution.
* Track configuration changes you have made since a specific baseline.
* Create clear document records of changes going into the production environment.
* Build confidence in getting things right when making changes to the deployed solution.

## Prerequisites:

1. .NET Framework 4.5 to be able to run the tool.
2. [FIMAutomation](https://docs.microsoft.com/en-us/powershell/module/fimautomation/?view=idm-ps-2016sp1) PowerShell snap-in must be installed locally to export and generate report for the MIM Service configuration.
3. A modern browser (e.g. Microsoft Edge) to view the report.

## Obtaining and testing the tool:

* Download the latest release of your choice from the [releases](https://github.com/Microsoft/MIMConfigDocumenter/releases) section under the Code tab in the project. We recommend downloading the MIMConfigDocumenter.zip file
* Be sure to *Unblock* the downloaded zip file before extracting the contents to an empty local folder. This process will extract the MIM Configuration Documenter application binaries, along with sample data files for the Contoso Corporation.
* Test the tool by executing the included PowerShell script InvokeDocumenter-Contoso.ps1.

## Using the tool:

To generate documentation, the MIM Configuration Documenter compares any provided point in time configuration export, with another. The difference between the two configurations is used to generate the report.

### Development to Production Changelog
A great use for the MIM Configuration Documenter is to track and highlight changes that are going to be transported from a staging/development environment into a production solution. For this scenario, an export of both the **Pre-production** / **Pilot** and **Production** environments can be compared with each other. This makes a great attachment to a change control or source code commit (if your configuration is exported and stored in source control).

### Deployment Documentation
As part of any new MIM , a good value add is to create a set of documents outlining the initial deployment. This can be an internal configuration record or a customer document as part of a consulting engagement. In this case (or other scenarios where there is only a single MIM environment) the MIM Configuration Documenter provides several baseline configurations that can be used as a baseline for document generation. These includes:

* FIM-Base_4.0.3684.2
* FIM-R2-SP1-Base_4.1.3419.0
* FIM-R2-SP1-Base_4.1.3461.0
* FIM-R2
* MIM-SP1-Base_4.4.1302.0
* MIM-SP1-Base_4.4.1459.0

 ## Step by step guide:

* Export the MIM Synchronization Server and if present the MIM Service configuration of your environment(s).
	* MIM Sync Server: In the MIM Synchronization console export the server configuration by using **File** | **Export Server Configuration** into an empty directory.
	* MIM Service Configuration: The MIM Service and Portal configuration can be exported using the **[ExportSchema.ps1](https://github.com/microsoft/MIMConfigDocumenter/blob/master/src/MIMConfigDocumenterCmd/Scripts/ExportSchema.ps1)** and **[ExportPolicy.ps1](https://github.com/microsoft/MIMConfigDocumenter/blob/master/src/MIMConfigDocumenterCmd/Scripts/ExportPolicy.ps1)** scripts located in the *[/scripts](https://github.com/microsoft/MIMConfigDocumenter/blob/master/src/MIMConfigDocumenterCmd/Scripts)* directory.
	* As previously mentioned, if you have configuration files from only a single environment, you can use any one set of the FIM or MIM config files provided with the tool as your **Production** / **Baseline** / **Reference** environment.
* Copy the configuration export files produced in the previous step to the */SyncConfig* and */ServiceConfig* folders under the */Data* directory of the MIM Configuration Documenter tool.
	* As an example, the **Pilot** configuration files for the customer **Contoso** are provided as a sample in the */Data\Contoso\Pilot\SyncConfig*" and */Data\Contoso\Pilot\ServiceConfig* directories.
	* **NOTE:** The names of the FIM/MIM Service schema and policy export files must be *schema.xml* and *policy.xml* respectively.
* Make a copy of the *InvokeDocumenter-Contoso.ps1* script, name it appropriately and then open and edit the new script using the instructions provided in the script.

	| Parameter        | Description |
	|------------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
	| pilotConfig      | Provide the path of the **Reference** / **Baseline** / **Pilot** configuration export files relative to the MIM Configuration Documenter */Data* folder.                                   |
	| productionConfig | Provide the path of the **Production** / **Baseline** configuration export files relative to the MIM Configuration Documenter */Data* folder.                                              |
	| reportType       |  Select the components to include in the document * **SyncOnly** for only a MIM Synchronization Engine * **ServiceOnly** for only the MIM Service and Portal * **SyncAndService** for both |

* Run your script to generate the documentation report of your configuration exports. Upon successful execution, the report will be placed in the /*Report* folder.

# Contributing

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
