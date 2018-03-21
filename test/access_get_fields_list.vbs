' ====================================================================
'
' Author : Christophe Avonture
' Date	: November 2017
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
' Requires
' ========
'
' * src\classes\MSAccess.vbs
' * src\classes\MSExcel.vbs
'
' To get more info, please read https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSAccess.md#getfieldslist
'
' Changes
' =======
'
' March 2018 - Open Excel and no more Notepad once finished
'
' ====================================================================

Option Explicit

Sub ShowHelp()

	wScript.echo " =========================================="
	wScript.echo " = Scan for fields in MS Access databases ="
	wScript.echo " =========================================="
	wScript.echo ""
	wScript.echo " Please specify the name of the database to scan; f.i. : "
	wScript.echo " " & Wscript.ScriptName & " 'C:\Temp\db1.accdb'"
	wScript.echo ""

	wScript.echo "To get more info, please read https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSAccess.md#getfieldslist"
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
	IncludeFile(sFolder & "src\classes\MSExcel.vbs")

End Sub

Dim cMSAccess, cMSExcel
Dim arrDBNames(0)
Dim sFieldsList, sFileName, sFile
Dim objFSO, objFile, oShell

	' Get the first argument (f.i. "C:\Temp\db1.accdb")
	If (wScript.Arguments.Count = 0) Then

		Call ShowHelp

	Else

		' Get the path specified on the command line
		sFile = Wscript.Arguments.Item(0)

		' Includes external classes
		Call IncludeClasses

		Set cMSAccess = New clsMSAccess

		cMSAccess.Verbose = True

		arrDBNames(0) = sFile

		' Get the list of fields for each table in the
		' specified databases
		sFieldsList = cMSAccess.GetFieldsList(arrDBNames)

		Set cMSAccess = Nothing

		Set objFSO = CreateObject("Scripting.FileSystemObject")

		' Finally, output the list into a flatfile and open it
		sFileName = objFSO.GetSpecialFolder(2) & "\output.csv"

		Set objFile = objFSO.CreateTextFile(sFileName, 2, True)
		objFile.Write sFieldsList
		objFile.Close
		Set objFile = Nothing

		Set cMSExcel = New clsMSExcel
		cMSExcel.FileName = sFileName
		cMSExcel.Verbose = True
		cMSExcel.OpenCSV sFile & " - Field lists", "fields"
		Call cMSExcel.MakeVisible
		Set cMSExcel = Nothing
	End if