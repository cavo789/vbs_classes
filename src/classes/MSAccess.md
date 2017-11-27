# MSAccess.vbs

## Decompose

Open a database and export every forms, macros, modules and reports code.

Contents will be saved in a /src folder.

### Sample script 

See also the test script in folder `/test` or online : https://github.com/cavo789/vbs_scripts/blob/master/test/access_decompose.vbs

```VB
Set cMSAccess = New clsMSAccess

cMSAccess.Verbose = True

arrDBNames(0) = "C:\Temp\db1.accdb"

' The second parameter is where sources files should be stored
' If not specified, will be in the same folder where the database is 
' stored, in the /src subfolder (will be created if needed)
Call cMSAccess.Decompose(arrDBNames, "")

Set cMSAccess = Nothing
```

## GetFieldsList

This function will scan every tables of a MS Access database and will generate a .csv file containing theses informations : 

| Filename | TableName | FieldName | FieldType | FieldSize | ShortestSize | LongestSize | Position | Occurences |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| xxx | xxx | xxx | xxx | xxx | xxx | xxx | xxx | xxx |

* `Filename` : Fullname of the database, absolute filepath on disk.
* `TableName` : The name of the table within the scanned database.
* `FieldType` : The type of the field (Autonumber, Double, Integer, Text, ...).
* `FieldName` : The name of the field in the table *(for instance 255 for a text field)* 
* `FieldSize` : The size of the field (as defined in the field's properties).
* `ShortestSize` : For text and memo fields, the lenght of the shortest information stored in that field *(for instance, the size will be 3 if the shortest value is `mdb`)*
* `LongestSize` : For text and memo fields, the lenght of the longest information stored in that field *(so if the longest size is 125 and the fieldsize is 255, we known that we can perhaps reduce the FieldSize to 125 and save space on the disk)* 
* `Position` : The position of that the field in the table structure
* `Occurences` : The number of times that the same table name and same field name has been found. When you're scanning a single database, `Occurrences` will always be set to 1 but when `GetFieldsList` is used to scan more than one databases, `Occurences` will allows you to retrieve fields defined in more than one database : the same table name and the same field name in three databases will return 3 f.i.

### Sample script 

See also the test script in folder `/test` or online : https://github.com/cavo789/vbs_scripts/blob/master/test/access_get_fields_list.vbs

```VB
Set cMSAccess = New clsMSAccess

cMSAccess.Verbose = True

arrDBNames(0) = "C:\Temp\db1.accdb"
arrDBNames(1) = "C:\Temp\db2.accdb"
arrDBNames(2) = "C:\Temp\db3.accdb"

' Get the list of fields for each table in the specified databases
sFieldsList = cMSAccess.GetFieldsList(arrDBNames)

Set cMSAccess = Nothing

wScript.echo sFieldList
```

## GetListOfTables

Get the list of tables of one or more MS Access databases

### Sample script 

See also the test script in folder `/test` or online : https://github.com/cavo789/vbs_scripts/blob/master/test/access_get_list_of_tables.vbs

```VB
Set cMSAccess = New clsMSAccess

arrDBNames(0) = "C:\Temp\db1.accdb"
arrDBNames(1) = "C:\Temp\db2.accdb"
arrDBNames(2) = "C:\Temp\db3.accdb"

wScript.Echo cMSAccess.GetListOfTables(arrDBNames, false)

Set cMSAccess = Nothing
```

## RemovePrefix

This script will scan the specified databases and will loop accross all tables : if the table name start with the specified prefix (f.i. `dbo_`), that prefix will be removed. So if a table is called `dbo_test`, after the execution of the RemovePrefix function, the table name will be `test`.

### Sample script 

See also the test script in folder `/test` or online : https://github.com/cavo789/vbs_scripts/blob/master/test/access_remove_prefix.vbs

```VB
Set cMSAccess = New clsMSAccess

cMSAccess.Verbose = True

arrDBNames(0) = "C:\Temp\db1.accdb"

Call cMSAccess.RemovePrefix(arrDBNames, "dbo_")

Set cMSAccess = Nothing
```