# MSSQL.vbs

This class provide functionnalities for working with a SQL Server database

## Table of content

- [CreateUserDSN](#createuserdsn)
	- [Sample script](#sample-script)
- [GetTableContent](#gettablecontent)
	- [Sample script](#sample-script)
- [GetTableContentMarkdown](#gettablecontentmarkdown)
	- [Sample script](#sample-script)

## CreateUserDSN

Create a User DSN to access a database through ODBC.

### Sample script

See https://github.com/cavo789/vbs_scripts/blob/master/test/sql_create_dsn.vbs for an example

```VB
Set cMSSQL = New clsMSSQL

cMSSQL.Verbose = True

cMSSQL.ServerName = "ServerName"
cMSSQL.DatabaseName = "DatabaseName"
cMSSQL.UserName = "UserName"

wScript.echo cMSSQL.CreateUserDSN(array("dsn DB"))

Set cMSSQL = Nothing
```

## GetTableContent

Read an entire table and generate a string with the table content. This string can f.i. be a CSV delimited string.

### Notes

1. You can choose the delimiter (can be ";", ",", "|" or something else); just initialize the Delimiter() property
2. Will probably give errors on big table since the result is a string

### Sample script

See https://github.com/cavo789/vbs_scripts/blob/master/test/sql_GetTableContent.vbs for an example

```VB
Set cMSSQL = New clsMSSQL

cMSSQL.Verbose = True

cMSSQL.ServerName = "srvName"
cMSSQL.DatabaseName = "dbName"
cMSSQL.Delimiter = ";"

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

See https://github.com/cavo789/vbs_scripts/blob/master/test/sql_GetTableContentMarkdown.vbs for an example

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