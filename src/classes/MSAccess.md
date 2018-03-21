# MSAccess.vbs

This script exposes a VB Script class for working with MS Access databases.

## Table of content

- [AttachTable](#attachtable)
- [CheckIfTableExists](#checkiftableexists)
- [Decompose](#decompose)
- [GetFieldsList](#getfieldslist)
- [GetListOfTables](#getlistoftables)
- [OpenDatabase](#opendatabase)
- [RemovePrefix](#removeprefix)

## AttachTable

Add a new attached-table in the MS Access database.

### Sample script
See https://github.com/cavo789/vbs_scripts/blob/master/test/access_attach_table.vbs for an example

```vbnet
Set cMSAccess = New clsMSAccess

cMSAccess.DatabaseName = "c:\temp\db1.accdb"
cMSAccess.OpenDatabase

If (Not(cMSAccess.CheckIfTableExists("users"))) Then

	' SQL Servername, DatabaseName, Source table name,
	' local table name and True for "Use trusted connection"
	cMSAccess.AttachTable("Server", "DBName", "dbo.users", _
		"users", True)

End If

cMSAccess.CloseDatabase

Set cMSAccess = Nothing
```
## CheckIfTableExists

Verify if a specific table exists in an existing database.

This function can be used before running a SQL statement, before adding a new attached table, ...

### Sample script
See https://github.com/cavo789/vbs_scripts/blob/master/test/access_check_table.vbs for an example

```vbnet
Set cMSAccess = New clsMSAccess

cMSAccess.DatabaseName = "c:\temp\db1.accdb"
cMSAccess.OpenDatabase

bExists = cMSAccess.CheckIfTableExists("users")

If (bExists) Then
	wScript.echo "The table 'users' is well in"
Else
	wScript.echo "There is no table called 'users' in the db"
End If

cMSAccess.CloseDatabase

Set cMSAccess = Nothing
```

## Decompose

Open a database and automate the extraction of `forms`, `macros`, `modules` and `reports` and export them as text files stored in a `/src` file where the database is stored.

### Sample script
See https://github.com/cavo789/vbs_scripts/blob/master/test/access_decompose.vbs for an example

```vbnet
Set cMSAccess = New clsMSAccess

cMSAccess.Verbose = True

arrDBNames(0) = "c:\temp\db1.accdb"

' The second parameter is where sources files should be stored
' If not specified, will be in the same folder where the database is
' stored, in the /src subfolder (will be created if needed)
Call cMSAccess.Decompose(arrDBNames, "")

Set cMSAccess = Nothing
```

By calling this code, you'll get, for instance :

```text
Process database c:\temp\db1.accdb
	Export module getFigures to c:\temp\db1_accdb\src\Modules\getFigures.txt
	Export module Helper to c:\temp\db1_accdb\src\Modules\Helper.txt
	Export module Constants to c:\temp\db1_accdb\src\Modules\Constants.txt
	Export module clsWindows to c:\temp\db1_accdb\src\Modules\clsWindows.txt
	Export module clsFolder to c:\temp\db1_accdb\src\Modules\clsFolder.txt
	Export module clsFile to c:\temp\db1_accdb\src\Modules\clsFile.txt
	Export macro getFigures to c:\temp\db1_accdb\src\Macros\getFigures.txt
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

See https://github.com/cavo789/vbs_scripts/blob/master/test/access_get_fields_list.vbs for an example

```vbnet
Set cMSAccess = New clsMSAccess

cMSAccess.Verbose = True

arrDBNames(0) = "c:\temp\db1.accdb"
arrDBNames(1) = "c:\temp\db2.accdb"
arrDBNames(2) = "c:\temp\db3.accdb"

' Get the list of fields for each table in the specified databases
' sFieldsList will be a string containing a .csv content
sFieldsList = cMSAccess.GetFieldsList(arrDBNames)

wScript.echo sFieldsList

Set cMSAccess = Nothing
```

By calling this code, you'll get, for instance :

```text
Database;TableName;FieldName;FieldType;FieldSize;ShortestSize;LongestSize;Position;Occurences
C:\Temp\db1.accdb;Bistel;RefDate;Date/Time;8;;1;1
C:\Temp\db1.accdb;Bistel;BudgetType;Byte;1;;2;1
C:\Temp\db1.accdb;Bistel;OrganicDivision;Text (fixed width);2;;3;1
C:\Temp\db1.accdb;Bistel;Program;Text (fixed width);1;;4;1
C:\Temp\db1.accdb;Bistel;Published;Yes/No;1;;5;1
C:\Temp\db1.accdb;Bistel;DescriptionDutch;Text;50;10;48;6;1
C:\Temp\db1.accdb;Bistel;DescriptionFrench;Text;50;0;50;7;1
C:\Temp\db1.accdb;Bistel;Article;Text;6;6;6;8;1
C:\Temp\db1.accdb;departements;bud;Text;255;2;2;1;1
C:\Temp\db2.accdb;Bistel;RefDate;Date/Time;8;;1;3
C:\Temp\db3.accdb;Bistel;RefDate;Date/Time;8;;1;3
```

## GetListOfTables

Get the list of tables of one or more MS Access databases.

### Sample script

See https://github.com/cavo789/vbs_scripts/blob/master/test/access_get_list_of_tables.vbs for an example

```vbnet
Set cMSAccess = New clsMSAccess

arrDBNames(0) = "c:\temp\db1.accdb"
arrDBNames(1) = "c:\temp\db2.accdb"
arrDBNames(2) = "c:\temp\db3.accdb"

' Get the list of tables in the specified databases
' sTablesList will be a string containing a .csv content
sTablesList = cMSAccess.GetListOfTables(arrDBNames, false)

wScript.echo sTablesList

Set cMSAccess = Nothing
```

## OpenDatabase

Open the MS Access interface and load a given database.

### Sample script

See https://github.com/cavo789/vbs_scripts/blob/master/test/access_open_database.vbs for an example

```vbnet
Set cMSAccess = New clsMSAccess

cMSAccess.DatabaseName = "c:\temp\db1.accdb"
cMSAccess.OpenDatabase

' Do something with the database

cMSAccess.CloseDatabase

Set cMSAccess = Nothing
```

## RemovePrefix

This script will scan the specified databases and will loop accross all tables : if the table name start with the specified prefix (f.i. `dbo_`), that prefix will be removed. So if a table is called `dbo_test`, after the execution of the RemovePrefix function, the table name will be `test`.

### Sample script

See https://github.com/cavo789/vbs_scripts/blob/master/test/access_remove_prefix.vbs for an example

```vbnet
Set cMSAccess = New clsMSAccess

cMSAccess.Verbose = True

arrDBNames(0) = "c:\temp\db1.accdb"

Call cMSAccess.RemovePrefix(arrDBNames, "dbo_")

Set cMSAccess = Nothing
```