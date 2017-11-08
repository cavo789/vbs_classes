' ===========================================================================
'
' Author : Christophe Avonture
' Date   : November 2017
'
' Scan a folder recursively for MS Access databases (.accdb or .mdb)
' and then, for each DB, retrieve the list of tables and get properties like,
' when the table is an attached one, the original DB name, 
' original table name, ...
'
' Once finished, Notepad will be started and will show the result : a .csv 
' file that can be used f.i. in Excel
'
' Requires 
' ========
'
' 	* src\classes\ScanFolder.vbs
' 	* src\classes\MSAccess.vbs
'
' ===========================================================================

Option Explicit

' Include the script library in this context
Sub IncludeFile(sFileName) 

	Dim objFSO, objFile

	Set objFSO = CreateObject("Scripting.FileSystemObject")    

	If (objFSO.FileExists(sFileName)) Then

		Set objFile = objFSO.OpenTextFile(sFileName, 1)  ' ForReading

		ExecuteGlobal objFile.ReadAll()

		objFile.close

	Else

		wScript.Echo "ERROR - IncludeFile - File " & sFileName & " not found!"

	End If

	Set objFSO = Nothing

End Sub

' Included needed classes
Sub IncludeClasses()

	' Get fullpath for the needed classes files, located in the parent folder
	' (this sample script is in the /src/test folder and the class is in 
	' the /src/classes folder)
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")		
	Set objFile = objFSO.GetFile(Wscript.ScriptName)
	sFolder = objFSO.GetParentFolderName(objFile) & "\"
	sFolder = objFSO.GetParentFolderName(sFolder) & "\"
	Set objFile = Nothing

	IncludeFile(sFolder & "src\classes\ScanFolder.vbs")
	IncludeFile(sFolder & "src\classes\MSAccess.vbs")
	
End Sub

Sub ShowHelp()

    Wscript.Echo " ====================================================="
    Wscript.Echo " = Scan for MS Access databases with attached tables ="
    Wscript.Echo " ====================================================="
    WScript.Echo ""    
    WScript.Echo " Please specify the name of the folder to scan; f.i. : "
    WScript.Echo " " & Wscript.ScriptName & " 'C:\FolderName'"
    WScript.Echo ""

    WScript.Quit 

End sub

Dim cScanFolder
Dim objFSO, objFile, oShell
Dim sFolder, sLines, sFileName, sTablesList
Dim arrDBNames

	' Get the first argument (f.i. "c:\temp\")
	If (wScript.Arguments.Count = 0) Then 

		Call ShowHelp

	Else 

		' Includes external classes
		Call IncludeClasses

		Set cScanFolder = New clsScanFolder

		' Create the ouput file in the temporary folder (2)
		' Will be something like C:\Users\username\AppData\Local\Temp\output.txt
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
		
			' cScanFolder has generated the output.txt file with the 
			' list of MS Access databases found in the search folder.
			' Process now these files

			Dim cMSAccess
			Set cMSAccess = New clsMSAccess

			' Get the output file content 
			Set objFile = objFSO.OpenTextFile(cScanFolder.OutputFileName, 1) 
			sLines = objFile.ReadAll()

			' Just in case : remove empty lines (i.e. two or more vbCrLf)
			sLines = Replace(sLines, vbCrLf & vbCrLf, vbCrLf)

			' Remove final vbCrLf
			If Right(sLines, 2) = vbCrLf Then sLines = Left(sLines, Len(sLines) - 2)

			' Now we've one line = one database name, convert into an array
			arrDBNames = Split(sLines, vbCrLf)  

			' And get the list of tables for these DBs
			sTablesList = cMSAccess.GetListOfTables(arrDBNames, false)

			Set cMSAccess = Nothing

			' Finally, output the list into a flatfile and open it
			sFileName = objFSO.GetSpecialFolder(2) & "\output.csv" 

			Set objFile = objFSO.CreateTextFile(sFileName, 2, True)
			objFile.Write sTablesList
			objFile.Close
			Set objFile = Nothing

			Set oShell = WScript.CreateObject ("WScript.Shell")
			oShell.run "notepad.exe """ & sFileName & ""
			Set oShell = Nothing

		End If

		Set cScanFolder = Nothing
		Set objFSO = Nothing

	End If