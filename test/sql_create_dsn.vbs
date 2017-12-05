' =======================================================================
'
' Author : Christophe Avonture
' Date	: December 2017
'
' Create a system DSN
'
' Requires
' ========
'
' * src\classes\MSSQL.vbs
'
'=======================================================================

Option Explicit

Const cServerName	= "" ' <-- Mandatory, specify your server name
Const cDatabaseName	= "" ' <-- Mandatory, specify your database name
Const cUserName		= "" ' <-- Mandatory, specify the user to use

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

	' Get fullpath for the needed classes files, located in the parent folder
	' (this sample script is in the /src/test folder and the class is in
	' the /src/classes folder)

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(Wscript.ScriptName)
	sFolder = objFSO.GetParentFolderName(objFile) & "\"
	sFolder = objFSO.GetParentFolderName(sFolder) & "\"
	Set objFile = Nothing

	IncludeFile(sFolder & "classes\MSSQL.vbs")

End Sub

Dim cMSSQL

	' Includes external classes
	Call IncludeClasses

	Set cMSSQL = New clsMSSQL

	cMSSQL.Verbose = True

	cMSSQL.ServerName = cServerName
	cMSSQL.DatabaseName = cDatabaseName
	cMSSQL.UserName = cUserName

	wScript.echo cMSSQL.CreateUserDSN(array(cDatabaseName))

	Set cMSSQL = Nothing
