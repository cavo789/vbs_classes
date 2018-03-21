# Debug.vbs

Helpers to help debugging a .vbs script

## Table of content

- [echo](#echo)

## echo

Output a text information on the console or by using a popup dialog (when started with wscript.exe)

### Sample script

```vbnet
Set cDebug = New clsDebug

cDebug.enable = true
cDebug.onlyConsole = true
cDebug.echo "You've started this script from the command line (cscript.exe)"

cDebug.onlyConsole = False
cDebug.echo "Even if started from Windows, you'll see this message"

Set cDebug = Nothing
```