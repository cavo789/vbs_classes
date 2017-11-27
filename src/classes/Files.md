# Files.vbs

This script exposes a VB Script class for working with files.

## CreateTextFile

### Sample script

```VB
Set cFiles = new clsFiles
cFiles.Verbose = True

cFiles.CreateTextFile "c:\temp\test.txt", "your content"
cFiles.CreateTextFile "c:\temp\test2.txt", "your second content"
cFiles.CreateTextFile "c:\temp\test3.txt", "your third content"

Set cFiles = Nothing
```