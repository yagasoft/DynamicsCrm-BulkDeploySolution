# DynamicsCrm-ExportSolution

(No longer maintained!)

### Version: 2.1.2
---

A command-line tool that can export multiple solutions and deploy them to multiple environments in one go. Supports manual retry of failed imports during the current run.

### Guide

Configuration is in the following format:

```json
{
	"sourceConnectionString": "AuthType=AD; Url=...",
	"destinationConnectionStrings": [
		"AuthType=AD; Url=...",
		"AuthType=AD; Url=..."
	],
	"solutionConfigs": [
		{
			"solutionName": "Solution1",
			"isManaged": false
		},
		{
			"solutionFile": "Solution2_0_0_0_1.zip"
		}
	]
}
```

The configuration filename must be passed as a command-line argument.

The following options can be passed to the command-line:

Parameter | Description
:---: | ---
-f | Pass a list of solution configuration files to process (space-separated).
-c | Pass a connections configuration file to use. Only used when no connection info can be found in the solution configuration file.
-P | Prevents the program from pausing on exit.
-r | Enables automatic import failure retry. Must be followed by an integer (min: 0).

## Changes

#### _v2.1.2 (2019-01-01)_
+ Improved: switched to EnhancedOrgService for better performance
+ Fixed: not handling error in imports
#### _v2.1.1 (2018-12-20)_
+ Added: progress indicator
+ Added: detailed errors and import log on failure
+ Added: optionally separated the connections from the solution configuration
+ Added: support for multiple solution configurations in one go
+ Added: support for regex in solution file names. Removes the need for editing the configuration everytime the exported solution version changes.
+ Changed: 'solutionPath' param is not split into 'solutionFolder' and 'solutionFile'. The folder can be relative to the executable file.
#### _v1.2.2 (2018-09-26)_
+ Added: additional command-line arguments
#### _v1.1.1 (2018-09-10)_
+ Initial release

---
**Copyright &copy; by Ahmed el-Sawalhy ([Yagasoft](http://yagasoft.com))** -- _GPL v3 Licence_
