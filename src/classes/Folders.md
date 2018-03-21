# Files.vbs

This script exposes a VB Script class for working with files.

## Table of content

- [CountLines](#countlines)
- [Create](#Create)
- [Exists](#exists)
- [getCurrentFolder](#getcurrentfolder)

## countLines

Get all files in the specifed folder and if the extension is the one specified, read the file and count the number of lines in that file.

It's an easy way to, f.i., count the number of lines in every .txt files of a folder

### Sample script

```vbnet
Dim sCurrentFolder
Dim objDict, objKey

Set cFolders = new clsFolders
sCurrentFolder = cFolders.getCurrentFolder()
Set objDict = cFolders.countLines(sCurrentFolder, "txt")

For Each objKey In objDict

	sFileName = objKey
	wCount = objDict(objKey)

	wScript.echo sFileName & " has " & wCount & " lines"

Next

Set objDict = Nothing
Set objKey = Nothing
Set cFolders = Nothing
```

## Create

Create a folder structure, create any parent folder if needed

See [../../test/folder_create.vbs](../../test/folder_create.vbs) for an example

## Exists

Check if a folder exists

### Sample script

```vbnet
Set cFolders = new clsFolders
If (cFolders.exists("C:\Temp\")) Then
	wScript.echo "Yes, the folder exists"
End If
Set cFolders = Nothing
```
## getCurrentFolder

Return the current folder i.e. the folder from where the script has been started

### Sample script

```vbnet
Dim sCurrentFolder
Set cFolders = new clsFolders
sCurrentFolder = cFolders.getCurrentFolder()
Set cFolders = Nothing
```