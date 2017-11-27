# MSAccess.vbs

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

## Sample script 

```VB
Set cMSAccess = New clsMSAccess

cMSAccess.Verbose = True

arrDBNames(0) = "C:\Temp\db1.accdb"

' Get the list of fields for each table 
sFieldsList = cMSAccess.GetFieldsList(arrDBNames)

Set cMSAccess = Nothing

wScript.echo sFieldList
```    