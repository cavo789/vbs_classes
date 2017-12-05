' =======================================================================
'
' Author : Christophe Avonture
' Date	: December 2017
'
' Connect to a SQL Server database and export a table content
' as a csv string. The delimiter can be set by using the
' Delimiter() property of the clsMSSQL class.
'
' Requires
' ========
'
' * src\classes\MSSQL.vbs
'
' More info and explanations : please read https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSSQL.md#
' =======================================================================

Option Explicit

Const cServerName	= "" '<-- You need to specify your server name
Const cDatabaseName = "" '<-- You need to specify your database name
Const cTableName 	= "" '<-- You need to specify your table name

Sub ShowHelp()

	wScript.echo " =================================="
	wScript.echo " = GetTableContent for SQL Server ="
	wScript.echo " =================================="
	wScript.echo ""
	wScript.echo " You need to specify, at least, "
	wScript.echo ""
	wScript.echo "	 * Your SQL Server name (f.i. srvMSSQL)"
	wScript.echo "	 * Your Database name (f.i. dbRepo)"
	wScript.echo "	 * A table name (f.i. dboTest)"
	wScript.echo ""
	wScript.echo " Please edit the script and set the three constants, "
	wScript.echo " see declaration at the top of the script. "

	wScript.quit

End sub

' Include the script library in this context
Sub IncludeFile(sFileName)

	Dim objFSO, objFile

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If (objFSO.FileExists(sFileName)) Then
		Set objFile = objFSO.OpenTextFile(sFileName, 1)  ' ForReading
		ExecuteGlobal objFile.ReadAll()
		objFile.close
	Else
		wScript.echo "ERROR - IncludeFile - File " & sFileName & " not found!"
	End If

	Set objFSO = Nothing

End Sub

' Included needed classes
Sub IncludeClasses()

	Dim objFSO, objFile
	DIm sFolder

	' Get fullpath for the needed classes files, located in the
	' parent folder
	' (this sample script is in the /src/test folder and the class is in
	' the /src/classes folder)

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(Wscript.ScriptName)

	sFolder = objFSO.GetParentFolderName(objFile) & "\"
	sFolder = objFSO.GetParentFolderName(sFolder) & "\"

	Set objFile = Nothing

	IncludeFile(sFolder & "src\classes\MSSQL.vbs")

End Sub

Dim cMSSQL

	If (cServerName="") or (cDatabaseName="") or (cTableName="") Then

		Call ShowHelp

	Else

		' Includes external classes
		Call IncludeClasses

		Set cMSSQL = New clsMSSQL

		cMSSQL.Verbose = True

		cMSSQL.ServerName = cServerName
		cMSSQL.DatabaseName = cDatabaseName
		cMSSQL.Delimiter = ";"

		wScript.echo cMSSQL.DSN()

		wScript.echo cMSSQL.GetTableContent(cTableName)

		Set cMSSQL = Nothing

	End if
