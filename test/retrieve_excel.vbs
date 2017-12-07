' =======================================================================
'
' Author : Christophe Avonture
' Date	: November 2017
'
' Retrieve all Excel (xls, xlam, xlst, xlsx) files in a given
' folder. Then start Notepad and display the list of files found.
'
' Requires
' ========
'
' * src\classes\ScanFolder.vbs
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

	' Get fullpath for the needed classes files, located in the parent
	' folder (this sample script is in the /src/test folder and the class
	' is in the /src/classes folder)

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(Wscript.ScriptName)
	sFolder = objFSO.GetParentFolderName(objFile) & "\"
	sFolder = objFSO.GetParentFolderName(sFolder) & "\"
	Set objFile = Nothing

	IncludeFile(sFolder & "src\classes\ScanFolder.vbs")

End Sub

	Dim objFSO, objFile, oShell
	DIm sFolder, sTempFolder
	Dim cScanFolder

	' Get the first argument (f.i. "c:\temp\source.csv")
	If (wScript.Arguments.Count = 0) Then

		Call ShowHelp

	Else

		' Includes external classes
		call IncludeClasses

		Set cScanFolder = New clsScanFolder

		' Specify which extensions are to be retrieved
		cScanFolder.AddExtensions("xls")
		cScanFolder.AddExtensions("xlam")
		cScanFolder.AddExtensions("xlsm")
		cScanFolder.AddExtensions("xlst")
		cScanFolder.AddExtensions("xlsx")

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		sTempFolder = objFSO.GetSpecialFolder(2) & "\"
		cScanFolder.OutputFileName = sTempFolder & "\output.txt"

		' Get the path specified on the command line
		cScanFolder.SearchFolder = Wscript.Arguments.Item(0)

		Call cScanFolder.Run()

		Set oShell = wScript.CreateObject("WScript.Shell")
		oShell.run "notepad.exe " & sTempFolder & "\output.txt"
		Set oShell = Nothing

	End if