' =======================================================================
'
' Author : Christophe Avonture
' Date	: January 2017
'
' Create a folder recursively i.e. create all the structure if
' parent folders are not found
'
' Requires
' ========
'
' * src\classes\folders.vbs
'
' =======================================================================

Option Explicit

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

	IncludeFile(sFolder & "src\classes\Folders.vbs")

End Sub

Dim cFolders
Dim sFile

	' Includes external classes
	Call IncludeClasses

	Set cFolders = New clsFolders
	cFolders.Verbose = True

	' Create the full hierarchy
	Call cFolders.create("C:\Temp\A\B\C\D\E")