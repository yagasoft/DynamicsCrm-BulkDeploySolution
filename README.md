# DynamicsCrm-ExportSolution
### Version: 1.2.2
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
			"sourcePath": "Solution2_0_0_0_1.zip"
		}
	]
}
```

The configuration filename must be passed as a command-line argument.

The following options can be passed to the command-line _after_ the configuration filename:

Parameter | Description
:---: | ---
-P | Prevents the program from pausing on exit.
-r | Enables automatic import failure retry. Must be followed by an integer (min: 0).

## Changes

#### _v1.2.2 (2018-09-26)_
+ Added: additional command-line arguments

#### _v1.1.1 (2018-09-10)_
+ Initial release

---
**Copyright &copy; by Ahmed el-Sawalhy ([Yagasoft](http://yagasoft.com))** -- _GPL v3 Licence_
