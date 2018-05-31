# Helper.vbs

Helpers for working with .vbs scripts

## Table of content

- [ForceCScriptExecution](#forcecscriptexecution)

## ForceCScriptExecution

When the user double-clic on a .vbs file (from Windows explorer f.i.) the running process will be `WScript.exe` while it's `CScript.exe` when the .vbs is started from the command prompt.

This subroutine will check if the script has been started with cscript and if not, will run the script again with cscript and terminate the `wscript` version. This is usefull when the script generate a lot of `wScript.echo` statements, easier to read in a command prompt.

### Sample script

Just put these three lines at the top of your script :

```vbnet
Set cHelper = New clsHelper
Call cHelper.ForceCScriptExecution()
Set cHelper = Nothing
```
