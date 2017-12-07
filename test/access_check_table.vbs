' =====================================================================
'
' Author : Christophe Avonture
' Date	: November/December 2017
'
' Check if a specific table exists in a MS Access database
'
' Requires
' ========
'
' * src\classes\MSAccess.vbs
'
' To get more info, please read https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSAccess.md#checkiftableexists
' =====================================================================

Option Explicit

Sub ShowHelp()

	wScript.echo " =============================================="
	wScript.echo " = CheckIfTableExists for MS Access databases ="
	wScript.echo " =============================================="
	wScript.echo ""
	wScript.echo " Please specify the name of the database to process; f.i. : "
	wScript.echo " " & wScript.ScriptName & " 'C:\Temp\db1.accdb'"
	wScript.echo ""
	wScript.echo "To get more info, please read https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSAccess.md#checkiftableexists"
	wScript.echo ""

	wScript.quit

End sub

' Include the script library in this context
Sub IncludeFile(sFileName)

	Dim objFSO, objFile

	Set objFSO = CreateObject("Scripting.FileSystemObject")

	If (objFSO.FileExists(sFileName)) Then

		Set objFile = objFSO.OpenTextFile(sFileName, 1) ' ForReading

		ExecuteGlobal objFile.ReadAll()

		objFile.close

	Else

		wScript.echo "ERROR - IncludeFile - File " & _
			sFileName & " not found!"

	End If

	Set objFSO = Nothing

End Sub

' Included needed classes
Sub IncludeClasses()

	Dim objFSO, objFile
	DIm sFolder

	' Get fullpath for the needed classes files, located in the parent
	' folder (this sample script is in the /src/test folder and the
	' class is in the /src/classes folder)

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(Wscript.ScriptName)
	sFolder = objFSO.GetParentFolderName(objFile) & "\"
	sFolder = objFSO.GetParentFolderName(sFolder) & "\"
	Set objFile = Nothing

	IncludeFile(sFolder & "src\classes\MSAccess.vbs")

End Sub

Dim cMSAccess
Dim sFile, bExists
Dim arrDBNames(0)

	' Get the first argument (f.i. "C:\Temp\db1.accdb")
	If (wScript.Arguments.Count = 0) Then

		Call ShowHelp

	Else

		' Get the path specified on the command line
		sFile = Wscript.Arguments.Item(0)

		' Includes external classes
		Call IncludeClasses

		Set cMSAccess = New clsMSAccess

		cMSAccess.DatabaseName = sFile
		cMSAccess.OpenDatabase

		bExists = cMSAccess.CheckIfTableExists("users")

		If (bExists) Then
			wScript.echo "The table users exists"
		Else
			wScript.echo "There is no table called 'users' in the db"
		End If

		cMSAccess.CloseDatabase

		Set cMSAccess = Nothing

	End if
