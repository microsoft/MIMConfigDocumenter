<# 
    ExportSchema.ps1
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.

    The purpose of this script is to export the current schema configuration.

    The script stores the configuration into file "schema.xml" in the current directory.
#>

param
(
    $SchemaFilePath = (Join-Path $PWD -ChildPath "schema.xml")
)

if(@(get-pssnapin | where-object {$_.Name -eq "FIMAutomation"} ).count -eq 0) {add-pssnapin FIMAutomation}

Write-Host "Exporting configuration objects from pilot."
# Please note that SynchronizationFilter Resources inform the FIM MA.
$schema = Export-FIMConfig -schemaConfig -customConfig "/SynchronizationFilter"
if ($schema -eq $null)
{
    Write-Error "Export did not successfully retrieve configuration from FIM.  Please review any error messages and ensure that the arguments to Export-FIMConfig are correct."
}
else
{
    Write-Host "Exported " $schema.Count " objects from pilot."
    $schema | ConvertFrom-FIMResource -file $SchemaFilePath
    Write-Host "Pilot file is saved as " $SchemaFilePath "."
    if($schema.Count -gt 0)
    {
        Write-Host "Export complete.  The next step is to run ExportPolicy.ps1."
    }
    else
    {
        Write-Warning "While export completed, there were no resources.  Please ensure that the arguments to Export-FIMConfig are correct." 
     }
}