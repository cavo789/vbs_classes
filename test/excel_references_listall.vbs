' =====================================================================
' Author : Christophe Avonture
' Date	: January 2018
'
' Start Excel, open a file and get the list of references used by
' the file

' Requires
' ========
'
' * src\classes\MSExcel.vbs
' =====================================================================

Option Explicit

Sub ShowHelp()

	wScript.echo " =================================="
	wScript.echo " = Excel - Get list of references ="
	wScript.echo " =================================="
	wScript.echo ""
	wScript.echo " Please specify the name of the file to open; f.i. : "
	wScript.echo " " & Wscript.ScriptName & " 'c:\temp\sample.xlsx'"
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

	IncludeFile(sFolder & "src\classes\MSExcel.vbs")

End Sub

Dim cMSExcel
Dim sFile

	' Get the first argument (f.i. "c:\temp\source.csv")
	If (wScript.Arguments.Count = 0) Then

		Call ShowHelp

	Else

		' Get the path specified on the command line
		sFile = Wscript.Arguments.Item(0)

		' Includes external classes
		Call IncludeClasses

		Set cMSExcel = New clsMSExcel
		cMSExcel.Verbose = True

		cMSExcel.Instantiate
		cMSExcel.FileName = sFile

		Call cMSExcel.Open(False)

		' Get the list of all references used by the file
		' But limit the list to only external reference
		wScript.echo ""
		wScript.echo " **************************************"
		wScript.echo " * Get the list of external ressources *"
		wScript.echo " **************************************"
		Call cMSExcel.References_ListAll(True, False)

		' Same but, this time, only .xlam files
		wScript.echo ""
		wScript.echo " **************************************"
		wScript.echo " ************* ONLY .XLAM *************"
		wScript.echo " **************************************"
		Call cMSExcel.References_ListAll(True, True)

		Call cMSExcel.CloseFile

		cMSExcel.Quit

		Set cMSExcel = Nothing

	End If