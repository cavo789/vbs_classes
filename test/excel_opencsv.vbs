' ===========================================================================
'
' Author : Christophe Avonture
' Date   : November 2017
'
' Start Excel and open a csv file

' Requires 
' ========
'
' 	* src\classes\MSExcel.vbs
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

	IncludeFile(sFolder & "src\classes\MSExcel.vbs")
	
End Sub

Dim cMSExcel
Dim sFileName

	' Includes external classes
	Call IncludeClasses
	
	Set cMSExcel = New clsMSExcel
	
	cMSExcel.Verbose = True
	
	sFileName = "C:\Users\Avonture_christophe\AppData\Local\Temp\output.csv"
	
	cMSExcel.OpenCSV(sFileName, "A nice title", "Tab name")	