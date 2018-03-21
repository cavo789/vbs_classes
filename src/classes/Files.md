# Files.vbs

This script exposes a VB Script class for working with files.

## Table of content

- [CreateTextFile](#createtextfile)
- [IsReadOnly](#isreadonly)
- [SetReadOnly](#setreadonly)

## CreateTextFile

Create a text file on the disk

### Sample script

```vbnet
Set cFiles = new clsFiles
cFiles.Verbose = True

cFiles.CreateTextFile "c:\temp\test.txt", "your content"
cFiles.CreateTextFile "c:\temp\test2.txt", "your second content"
cFiles.CreateTextFile "c:\temp\test3.txt", "your third content"

Set cFiles = Nothing
```

## IsReadOnly

Return true if the specified file has the ReadOnly attribute set.

### Sample script

```vbnet
Set cFiles = new clsFiles
wScript.echo cFiles.IsReadOnly("c:\temp\test.txt")
Set cFiles = Nothing
```

## SetReadOnly

Set the ReadOnly attribute for a file. Assuming that the file exists.

### Sample script

```vbnet
Set cFiles = new clsFiles
cFiles.SetReadOnly("c:\temp\test.txt")
Set cFiles = Nothing
```