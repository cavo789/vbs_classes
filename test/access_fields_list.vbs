' ===========================================================================
'
' Author : Christophe Avonture
' Date   : November 2017
'
' Open a database, get the list of tables and for each of them, 
' get the list of fields and a few properties like the type, the size, 
' the shortest and longest value size (for text and memo fields)
'
' The output will be something like : 
' Database;TableName;FieldName;FieldType;FieldSize;ShortestSize;LongestSize;Position;Occurences;
' C:\Temp\db1.accdb;Bistel;RefDate;Date/Time;8;;1;1
' C:\Temp\db1.accdb;Bistel;BudgetType;Byte;1;;2;1
' C:\Temp\db1.accdb;Bistel;OrganicDivision;Text (fixed width);2;;3;1
' C:\Temp\db1.accdb;Bistel;Program;Text (fixed width);1;;4;1
' C:\Temp\db1.accdb;Bistel;Published;Yes/No;1;;5;1
' C:\Temp\db1.accdb;Bistel;DescriptionDutch;Text;50;10;48;6;1
' C:\Temp\db1.accdb;Bistel;DescriptionFrench;Text;50;0;50;7;1
' C:\Temp\db1.accdb;Bistel;Article;Text;6;6;6;8;1
' C:\Temp\db1.accdb;departements;bud;Text;255;2;2;1;1
'
' More info and explanations of fields : please read   https://github.com/cavo789/vbs_scripts/tree/master/src/classes/msaccess.md
'
' Requires 
' ========
'
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

	IncludeFile(sFolder & "src\classes\MSAccess.vbs")
	
End Sub

Dim cMSAccess
Dim arrDBNames(0)
Dim sFieldsList, sFileName
Dim objFSO, objFile, oShell

	' Includes external classes
	Call IncludeClasses
	
	Set cMSAccess = New clsMSAccess
	
	cMSAccess.Verbose = True
	
	arrDBNames(0) = "C:\Temp\db1.accdb"
	
	' Get the list of fields for each table 
	sFieldsList = cMSAccess.GetFieldsList(arrDBNames)

	Set cMSAccess = Nothing
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")		
	
	' Finally, output the list into a flatfile and open it
	sFileName = objFSO.GetSpecialFolder(2) & "\output.csv" 

	Set objFile = objFSO.CreateTextFile(sFileName, 2, True)
	objFile.Write sFieldsList
	objFile.Close
	Set objFile = Nothing

	Set oShell = WScript.CreateObject ("WScript.Shell")
	oShell.run "notepad.exe """ & sFileName & ""
	Set oShell = Nothing