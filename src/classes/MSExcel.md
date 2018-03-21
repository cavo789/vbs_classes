# MSExcel.vbs

This script exposes a VB Script class for working with MS Excel workbooks.

## Table of content

- [getApplicationLanguage](#getapplicationlanguage)
- [IsLoaded](#isloaded)
- [OpenCSV](#opencsv)
- [References_AddFromFile](#references_addfromfile)
- [References_ListAll](#references_listall)
- [References_Remove](#references_remove)
- [Save](#save)

## getApplicationLanguage

Retrieve MS Office application settings and extract the language used for the installation i.e. return the Excel's interface language.

Note : this function will only return FR (for French) or NL (for Dutch) since there are a lot of "IDs" (see https://technet.microsoft.com/en-us/library/cc179219.aspx#Language identifiers for the full list)

### Sample script

```vbnet
Set cMSExcel = New clsMSExcel
cMSExcel.Verbose = True

' Will return FR or NL; FR will be the default
wScript.echo cMSExcel.getApplicationLanguage("FR")
Set cMSExcel = Nothing
```

## IsLoaded

Check if a specific file is already opened in Excel
This function will return True if the file is already loaded.

### Sample script

```vbnet
Set cMSExcel = New clsMSExcel
If (cMSExcel.IsLoaded("test.xlsx")) then
	wScript.echo "Test file already loaded"
End if
Set cMSExcel = Nothing
```

## OpenCSV

Open a CSV file and convert it to a .xlsx file by adding a small template, title, auto-filter, FreezePanes, ...

### Sample script

```vbnet
Set cMSExcel = New clsMSExcel
cMSExcel.Verbose = True
cMSExcel.OpenCSV "c:\temp\source.csv", "My amazing title", "Tab name"
Set cMSExcel = Nothing
```

## References_ListAll

Display the list of references used by the file

The list can be limited to only external references and, too, to only .xlam files

See [../../test/excel_references_listall.vbs](../../test/excel_references_listall.vbs) for an example

## References_Remove

Remove a reference used by the file. Usefull for, f.i., removing a .xlam file used as a reference in the VBE editor

See [../../test/excel_references_remove.vbs](../../test/excel_references_remove.vbs) for an example

## References_AddFromFile

Add a reference (f.i. to a `.xlam` file) in an existing workbook.

See [../../test/excel_references_addfromfile.vbs](../../test/excel_references_addfromfile.vbs) for an example

## Save

Save the active sheet on disk

### Sample script

```vbnet
Set cMSExcel = New clsMSExcel
cMSExcel.Verbose = True
cMSExcel.OpenCSV "c:\temp\source.csv", "My amazing title", "Tab name"
cMSExcel.Save "c:\temp\source.xlsx"
Set cMSExcel = Nothing
```