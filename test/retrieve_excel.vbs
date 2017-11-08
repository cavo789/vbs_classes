' ===========================================================================
'
' Author : Christophe Avonture
' Date   : November 2017
'
' Retrieve all Excel (xls, xlam, xlst, xlsx) files in a given 
' folder. Then start Notepad and display the list of files found.
'
' Requires 
' ========
'
' 	* src\classes\ScanFolder.vbs
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
	
End Sub

	Dim objFSO, objFile, oShell
	DIm sFolder
	Dim cScanFolder
	
	' Includes external classes
	call IncludeClasses
	
	Set cScanFolder = New clsScanFolder
	
	' Specify which extensions are to be retrieved
	cScanFolder.AddExtensions("xls")
	cScanFolder.AddExtensions("xlam")
	cScanFolder.AddExtensions("xlsm")
	cScanFolder.AddExtensions("xlst")
	cScanFolder.AddExtensions("xlsx")
	cScanFolder.OutputFileName = "c:\temp\output.txt"
	cScanFolder.SearchFolder = "c:\temp"

	Call cScanFolder.Run()
    
	Set oShell = WScript.CreateObject ("WScript.Shell")
	oShell.run "notepad.exe c:\temp\output.txt"
	Set oShell = Nothing