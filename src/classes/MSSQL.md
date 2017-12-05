# MSSQL.vbs

This class provide functionnalities for working with a SQL Server database

## GetTableContent

Read an entire table and generate a string with the table content. This string can f.i. be a CSV delimited string.

### Notes

1. You can choose the delimiter (can be ";", ",", "|" or something else); just initialize the Delimiter() property
2. Will probably give errors on big table since the result is a string

### Sample script

See also the test script in folder `/test` or online : https://github.com/cavo789/vbs_scripts/blob/master/test/sql_GetTableContent.vbs

```VB
Set cMSSQL = New clsMSSQL

cMSSQL.Verbose = True

cMSSQL.ServerName = "srvName"
cMSSQL.DatabaseName = "dbName"

wScript.echo cMSSQL.GetTableContent("dbo.Test")

Set cMSSQL = Nothing
```

By calling this code, you'll get, for instance :

```text
fldname1;fldname2;fldname3
rec1_value1;rec1_value2;rec1_value3
rec2_value1;rec2_value2;rec2_value3
```
## GetTableContentMarkdown

Read an entire table and generate a string with the table content. Respect the markdown format.

### Notes

1. The delimiter will be "|" since the output should be a markdown string
2. Will probably give errors on big table since the result is a string

### Sample script

See also the test script in folder `/test` or online : https://github.com/cavo789/vbs_scripts/blob/master/test/sql_GetTableContentMarkdown.vbs

```VB
Set cMSSQL = New clsMSSQL

cMSSQL.Verbose = True

cMSSQL.ServerName = "srvName"
cMSSQL.DatabaseName = "dbName"

wScript.echo cMSSQL.GetTableContentMarkdown("dbo.Test")

Set cMSSQL = Nothing
```

By calling this code, you'll get, for instance :

```text
| fldname1 | fldname2 | fldname3 |
| rec1_value1 | rec1_value2 | rec1_value3 |
| rec2_value1 | rec2_value2 | rec2_value3 |
```