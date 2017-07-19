## MIM Configuration Documenter Change Log

All notable changes to MIMConfigDocumenter project will be documented in this file. The "Unreleased" section at the top is for keeping track of important changes that might make it to upcoming releases.

------------

### Version 1.17.0719.0

#### Fixed

* In the MIMWAL activity configuration, Value Expressions information gets reordered alphabetically.

------------

### Version 1.17.0610.0

#### Changed

* In the sync engine configuration, Inbound Scoping Filter information is displayed when documenting import attribute flows.
* In the sync engine configuration, sync rule scoping information is printed for join and projection rules.

#### Fixed

* In the sync engine configuration, import flows may get duplicated when inbound scoping filter in use.
* In the Sync Rule configuration, Inbound Scoping Filter configuration gets printed in Outbound Scoping Filter section
* In the MIMWAL activity configuration, Query Resources and Value Expressions information gets mixed up.

------------

### Version 1.17.0606.0

#### Fixed

* Tool crashing when MIMWAL RunPowerShellScript activity is configured with no input parameters.
* MV Object Deletion Rules configured not getting documented.
* Dereferencing of reference guids resuting in false positive changes being reported.

------------

### Version 1.17.0522.0 (Public Beta)

#### Added

* Baseline version check-in. Not all management agents supported by MIM / FIM are fully supported, so their documentation may be incomplete.

------------
