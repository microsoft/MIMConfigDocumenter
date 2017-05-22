<#
    ExportPolicy.ps1
	Copyright (c) Microsoft. All Rights Reserved.
	Licensed under the MIT license. See LICENSE.txt file in the project root for full license information.

    The purpose of this script is to export the current policy and synchronization configuration.

    The script stores the configuration into file "policy.xml" in the current directory.
#>

param
(
    $PolicyFilePath = (Join-Path $PWD -ChildPath "policy.xml")
)

if(@(get-pssnapin | where-object {$_.Name -eq "FIMAutomation"} ).count -eq 0) {add-pssnapin FIMAutomation}

Write-Host "Exporting configuration objects from pilot."
# In many production environments, some Set resources are larger than the default message size of 10 MB.
$policy = Export-FIMConfig -policyConfig -portalConfig -MessageSize 9999999
if ($policy -eq $null)
{
    Write-Error "Export did not successfully retrieve configuration from FIM.  Please review any error messages and ensure that the arguments to Export-FIMConfig are correct."
}
else
{
    Write-Host "Exported " $policy.Count " objects from pilot."
    $policy | ConvertFrom-FIMResource -file $PolicyFilePath
    Write-Host "Pilot file is saved as " $PolicyFilePath "."
    if($policy.Count -gt 0)
    {
        Write-Host "Export complete.  The next step is run SyncPolicy.ps1."
    }
    else
    {
        Write-Warning "While export completed, there were no resources.  Please ensure that the arguments to Export-FIMConfig are correct."
    }
}