' ======================================================================
'
' Author : Christophe Avonture
' Date	: November 2017
'
' Scan a folder recursively for MS Access databases (.accdb or .mdb).
'
' Once finished, Notepad will be started and will show the result : a .csv
' file with the list of files
'
' Requires
' ========
'
' * src\classes\ScanFolder.vbs
'
' ======================================================================

Option Explicit

Sub ShowHelp()

	wScript.echo " ================================"
	wScript.echo " = Scan for MS Access databases ="
	wScript.echo " ================================"
	wScript.echo ""
	wScript.echo " Please specify the name of the folder to scan; f.i. : "
	wScript.echo " " & Wscript.ScriptName & " 'c:\foldername'"
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

	' Get fullpath for the needed classes files, located in the parent
	' folder (this sample script is in the /src/test folder and the
	' class is in the /src/classes folder)

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.GetFile(Wscript.ScriptName)
	sFolder = objFSO.GetParentFolderName(objFile) & "\"
	sFolder = objFSO.GetParentFolderName(sFolder) & "\"
	Set objFile = Nothing

	IncludeFile(sFolder & "src\classes\ScanFolder.vbs")

End Sub

Dim cScanFolder
Dim objFSO, oShell
Dim sFolder

	' Get the first argument (f.i. "c:\temp\")
	If (wScript.Arguments.Count = 0) Then

		Call ShowHelp

	Else

		' Includes external classes
		Call IncludeClasses

		Set cScanFolder = New clsScanFolder

		' Create the ouput file in the temporary folder (2)
		' Will be something like
		' C:\Users\username\AppData\Local\Temp\output.txt
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		sFolder = objFSO.GetSpecialFolder(2) & "\"
		cScanFolder.OutputFileName = sFolder & "output.txt"

		' Define the list of extensions to retrieve (there are several
		' now for MS Access)
		cScanFolder.AddExtensions("accdb")
		cScanFolder.AddExtensions("adp")
		cScanFolder.AddExtensions("mdb")

		' Get the path specified on the command line
		cScanFolder.SearchFolder = Wscript.Arguments.Item(0)

		If cScanFolder.Run() Then

			Set oShell = WScript.CreateObject ("WScript.Shell")
			oShell.run "notepad.exe """ & cScanFolder.OutputFileName & ""
			Set oShell = Nothing

		End If

		Set cScanFolder = Nothing
		Set objFSO = Nothing

	End If