<#
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.

	This is a library of PowerShell cmdlets that provide a commmand-line interface to the MIM Configuration Documenter. 
#>

<#
.Synopsis
	Generates an HTML Report from the specified FIM / MIM configuration exports.
.DESCRIPTION
	Generates an HTML Report from the specified FIM / MIM configuration exports.
.PARAMETER PilotConfig
	The directory where the configuration exports of "pilot" / to-be system are located.
.PARAMETER ProductionConfig
	The directory where the configuration exports of "production" / as-is system are located.
.PARAMETER ReportType
	Optional. The type of report to generate. Possible values are: "SyncOnly", "ServiceOnly" or "SyncAndService".
	The default is "SyncAndService".
.EXAMPLE
	$pilotConfig = "Contoso\Pilot"
	$productionConfig = "Contoso\Base.R2-SP1_4.1.3419.0"
	$reportType = "SyncOnly"

	Publish-FIMSyncConfigReport -PilotConfig $pilotConfig -ProductionConfig $productionConfig -ReportType $reportType
#>
function Get-MIMConfigReport
{
	[CmdletBinding()]
	[OutputType([int])]
	Param
	(
		# Data Directory where Pilot configuration files are stored
		[Parameter(Mandatory = $true, Position = 0)]
		[string]
		$PilotConfig,
		
		# Data Directory where Production configuration files are stored
		[Parameter(Mandatory = $true, Position = 1)]
		[string]
		$ProductionConfig,

		#
        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateSet("SyncOnly","ServiceOnly","SyncAndService")]
        $ReportType = "SyncAndService"
	)

	Begin
	{
	}

	Process
	{
		if ($ReportType -ne "SyncOnly")
		{
            Compare-ServiceConfig -PilotConfig $PilotConfig -ProductionConfig $ProductionConfig -ExportType "Schema"
            Compare-ServiceConfig -PilotConfig $PilotConfig -ProductionConfig $ProductionConfig -ExportType "Policy"
		}

        Write-Verbose "Executing MIMConfigDocumenterCmd.exe"

		Start-Process `
		-FilePath 'MIMConfigDocumenterCmd.exe' `
		-WorkingDirectory $PWD `
        -ArgumentList "$PilotConfig $ProductionConfig $ReportType" `
        -Wait
    }

	End
	{
		Write-Verbose "Processed Get-FIMSyncConfigReport."
    }
}


<#
.Synopsis
    Compares the FIM Service configuration exports between the Pilot and Production system and generates the corresponding changes.xml file.
.DESCRIPTION
    Compares the FIM Service configuration exports between the Pilot and Production system and generates the corresponding changes.xml file.
.PARAMETER PilotConfig
	The directory where the configuration exports of "pilot" / to-be system are located.
.PARAMETER ProductionConfig
	The directory where the configuration exports of "production" / as-is system are located.
.PARAMETER ExportType
    Schema - Denotes that the export is for FIM Service Schema configuration.
    Policy - Denotes that the export is for FIM Service Policy configuration.
.PARAMETER PersonJoinCriteria
    The join criteria to be used to determine changes to the Person resources.
    The default is AccountName.
.PARAMETER GroupJoinCriteria
    The join criteria to be used to determine changes to the Group resources.
    The default is AccountName.
.EXAMPLE
	$pilotConfig = "Contoso\Pilot"
	$productionConfig = "Contoso\Base.R2-SP1_4.1.3419.0"
	$exportType = "Policy" # "Schema

	Compare-ServiceConfig -PilotConfig $pilotConfig -ProductionConfig $productionConfig -ExportType "Policy"
#>
function Compare-ServiceConfig
{
	[CmdletBinding()]
	[OutputType([int])]
	Param
	(
		# Data Directory where Pilot configuration files are stored
		[Parameter(Mandatory = $true, Position = 0)]
		[string]
		$PilotConfig,
		
		# Data Directory where Production configuration files are stored
		[Parameter(Mandatory = $true, Position = 1)]
		[string]
		$ProductionConfig,

		# Type of FIM Service config export
        [Parameter(Mandatory = $false, Position = 2)]
        [ValidateSet("Schema","Policy")]
		$ExportType,

        [Parameter(Mandatory = $false, Position = 3)]
        [string]
		$PersonJoinCriteria = "AccountName",

        [Parameter(Mandatory = $false, Position = 4)]
        [string]
		$GroupJoinCriteria = "AccountName"
    )

    Begin
    {
        if (@(Get-PSSnapin | Where-Object {$_.Name -eq "FIMAutomation"} ).Count -eq 0) { Add-PSSnapin "FIMAutomation" -ErrorAction Stop }
    }

    Process
    {
        $pilot_filename = "{0}\Data\{1}\ServiceConfig\{2}.xml" -f $PWD,$PilotConfig, $ExportType
        $production_filename = "{0}\Data\{1}\ServiceConfig\{2}.xml" -f $PWD, $ProductionConfig, $ExportType
        $changes_filename = "{0}\Data\Changes\ServiceConfig\{1}_AppliedTo_{2}_changes_{3}.xml" -f $PWD, $PilotConfig.Replace("\", "_"), $ProductionConfig.Replace("\", "_"), $ExportType

        if ((Test-Path ($changes_filename))) { Remove-Item $changes_filename }
		
		$changesDir = Split-Path $changes_filename -Parent
		if (!(Test-Path ($changesDir))) { New-Item $changesDir -ItemType Directory | Out-Null }

        if ($ExportType -eq "Policy")
        {
            Update-PolicyExport -PolicyFile $pilot_filename
            Update-PolicyExport -PolicyFile $production_filename
        }

        switch ($ExportType)
        {
            "Policy"
            {
                $JoinRules = @{
                    #=== Customer-dependent join rules ===
                    #Person and Group objects are not configuration will not be migrated.
                    #However, some configuration objects like Sets may refer to these objects.
                    #For this reason, we need to know how to join Person objects between
                    #systems so that configuration objects have the same semantic meaning.
                    Person = $PersonJoinCriteria;
                    Group = $GroupJoinCriteria;
    
                    #=== Policy configuration ===
                    #Sets, MPRs, Workflow Definitions, and so on. are best identified by DisplayName
                    #DisplayName is set as the default join criteria and applied to all object
                    #types not listed here.
    
                    #=== Schema configuration ===
                    #This is based on the system names of attributes and objects
                    #Notice that BindingDescription is joined using its reference attributes.
                    ObjectTypeDescription = "Name";
                    AttributeTypeDescription = "Name";
                    BindingDescription = "BoundObjectType BoundAttributeType";
    
                    #=== Portal configuration ===
                    ConstantSpecifier = "BoundObjectType BoundAttributeType ConstantValueKey";
                    SearchScopeConfiguration = "DisplayName SearchScopeResultObjectType Order";
                    ObjectVisualizationConfiguration = "DisplayName AppliesToCreate AppliesToEdit AppliesToView"
                }
            }

            "Schema"
            {

                $JoinRules = @{
                    #=== Schema configuration ===
                    #This is based on the system names of attributes and objects
                    #Notice that BindingDescription is joined using its reference attributes.
                    ObjectTypeDescription = "Name";
                    AttributeTypeDescription = "Name";
                    BindingDescription = "BoundObjectType BoundAttributeType";
                }
            }
        }

        Write-Verbose "Loading pilot file $pilot_filename."

        $pilot = ConvertTo-FIMResource -file $pilot_filename
        if($pilot -eq $null)
        {
            throw (New-Object NullReferenceException -ArgumentList "Check that the pilot config file '$pilot_filename' has data.")
        }

        $pilotCount = $pilot.Count

        Write-Verbose "`t`t$pilotCount pilot objects loaded."

        $production = ConvertTo-FIMResource -file $production_filename

        if($production -eq $null)
        {
            throw (New-Object NullReferenceException -ArgumentList "Check that the production config file '$production_filename' has data.")
        }

        $productionCount = $production.Count
        Write-Verbose "`t`t$productionCount production objects loaded."
       
        Write-Verbose "Executing join between pilot and production."
         
        $matches = Join-FIMConfig -Source $pilot -Target $production -Join $joinrules -DefaultJoin DisplayName

        if($matches -eq $null)
        {
            throw (New-Object NullReferenceException -ArgumentList "Matches is null. Check that the join succeeded and join criteria is correct for your environment.")
        }

        Write-Verbose "Executing compare between matched objects in pilot and production."
        $changes = $matches | Compare-FIMConfig

        if($changes -eq $null)
        {
            Write-Warning "$ExportType changes are null.  Check that no errors occurred while generating changes, or the systems compared are in sync and there are no changes to report."
        }

        $changesCount = $changes.count
        if ($changesCount -gt 0)
        {
            Write-Verbose "`t`tIdentified $changesCount changes to apply to production."
            Write-Verbose "Saving changes to $changes_filename."
            $changes | ConvertFrom-FIMResource -file $changes_filename
        }
        else
        {
            Write-Verbose "`t`tIdentified 0 changes to apply to production.  No Changes file will be created for report since systems are the same"
        }

        Write-Verbose "Compare Complete!"
    }

    End
    {
        Remove-PSSnapin "FIMAutomation" -ErrorAction SilentlyContinue
    }
}

<#
.Synopsis
   Updates Policy Export Files to Populate AccountName on ILMSync user and any other Users and Groups with missing AccountName.
   It also deletes the duplicate 'Users can create registration objects for themselves' MPR.
.DESCRIPTION
   Updates Policy Export Files to Populate AccountName on ILMSync user and any other Users and Groups with missing AccountName 
   so that Join-FIMConfig will succeed on default join criteria on AccountName.
   It also deletes the duplicate 'Users can create registration objects for themselves' MPR.
.EXAMPLE
   Update-PolicyExport policy.xml
.EXAMPLE
   Update-PolicyExport -PolicyFile policy.xml
#>
function Update-PolicyExport
{
    Param
    (
        [string] $PolicyFile = $(throw "PolicyFile parameter is required")
    )

    if (!(Test-Path($PolicyFile))) { throw "Policy xml file $PolicyFile not found" }

    $xmlUpdated = $false

	$policyXml = [xml] (Get-Content $PolicyFile)
	$xPathClause = "count(ResourceManagementObject//ResourceManagementAttribute[AttributeName = 'AccountName']) = 0"
	$personOrGroupWithoutAccountNameNodes = $policyXml.SelectNodes("//ExportObject[ResourceManagementObject[ObjectType = 'Person' or ObjectType = 'Group'] and $xPathClause]")

    foreach ($personNode in $personOrGroupWithoutAccountNameNodes)
    {
        $maikNicknameNode = $personNode.SelectSingleNode("self::node()//ResourceManagementAttribute[AttributeName = 'MailNickname' and count(Value) > 0]")

        if ($maikNicknameNode)
        { 
            $accountNameNode = $maikNicknameNode.CloneNode($true)
            $accountNameNode.SelectSingleNode("AttributeName").InnerText = "AccountName"
            [void]$maikNicknameNode.SelectSingleNode("parent::node()").InsertAfter($accountNameNode, $maikNicknameNode)
            $xmlUpdated = $true
        }
        else
        {
            $objectIDNode = $personNode.SelectSingleNode("self::node()//ResourceManagementAttribute[AttributeName = 'ObjectID' and count(Value) > 0]")
            $accountNameNode = $objectIDNode.CloneNode($true)
            $accountNameNode.SelectSingleNode("AttributeName").InnerText = "AccountName"
            [void]$objectIDNode.SelectSingleNode("parent::node()").InsertAfter($accountNameNode, $objectIDNode)
            $xmlUpdated = $true
        }
    }


	$xPathClause = "ResourceManagementObject//ResourceManagementAttribute[AttributeName = 'DisplayName' and Value = 'Users can create registration objects for themselves']"
	$xPathClause2 = "ResourceManagementObject//ResourceManagementAttribute[AttributeName = 'ActionType' and count(Values/string) > 1]"
	$mprNode = $policyXml.SelectSingleNode("//ExportObject[ResourceManagementObject[ObjectType = 'ManagementPolicyRule'] and $xPathClause and $xPathClause2]")

    if ($mprNode)
    {
        [void]$mprNode.SelectSingleNode("parent::node()").RemoveChild($mprNode)
        $xmlUpdated = $true
    }

    if ($xmlUpdated)
    {
        $policyXml.Save($PolicyFile)
    }
}

Export-ModuleMember -Function Get-MIMConfigReport

