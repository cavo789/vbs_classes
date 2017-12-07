# MSExcel.vbs

This script exposes a VB Script class for working with MS Excel workbooks.

## Table of content

- [OpenCSV](#opencsv)
	- [Sample script](#sample-script)
- [Save](#save)
	- [Sample script](#sample-script)

## OpenCSV

Open a CSV file and convert it to a .xlsx file by adding a small template, title, auto-filter, FreezePanes, ...

### Sample script

```VB
Set cMSExcel = New clsMSExcel
cMSExcel.Verbose = True
cMSExcel.OpenCSV "c:\temp\source.csv", "My amazing title", "Tab name"
Set cMSExcel = Nothing
```

## Save

Save the active sheet on disk

### Sample script

```VB
Set cMSExcel = New clsMSExcel
cMSExcel.Verbose = True
cMSExcel.OpenCSV "c:\temp\source.csv", "My amazing title", "Tab name"
cMSExcel.Save "c:\temp\source.xlsx"
Set cMSExcel = Nothing
```