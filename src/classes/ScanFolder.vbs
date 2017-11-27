' ===========================================================================
'
' Author : Christophe Avonture
' Date   : November 2017
'
' This script contains a VBS class to scan a folder structure and retrieve
' files with specific extensions (like, for instance : xls, xlsx, xlsm).
'
' The list of files will be outputted in a text file so that list can be 
' used in a separate process
'
' ===========================================================================

Option Explicit

Class clsScanFolder

	Private objFSO, objOutput
	Private sOutputFileName
	Private colExtensions
	Private sSearchFolder
	Private bVerbose 

	Public Property Let verbose(bYesNo)
		bVerbose = bYesNo
	End Property

	Private Sub Class_Initialize()	
		bVerbose = False
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set colExtensions = CreateObject("Scripting.Dictionary")
		colExtensions.CompareMode = 1 ' make lookups case-insensitive
	End Sub

	Private Sub Class_Terminate()
		objOutput.Close
		Set objOutput = Nothing
		Set colExtensions = Nothing
		Set objFSO = Nothing
	End Sub

	' Filenames will be outputted in this file
	Public Property Let OutputFileName(sFileName)
		sOutputFileName = sFileName
		Set objOutput = objFSO.OpenTextFile(sOutputFileName, 2, True) ' Writing
	End Property

	Public Property Get OutputFileName()
		OutputFileName = sOutputFileName
	End Property

	' Folder where to start the search
	Public Property Let SearchFolder(sFolderName)
		sSearchFolder = sFolderName
	End Property

	' ------------------------------------------------------------------
	'
	' List of extensions to search
	'
	' Example : 
	'
	'		' Define the list of extensions to retrieve
	'		Set cScanFolder = New clsScanFolder
	'		cScanFolder.AddExtensions("accdb")
	'		cScanFolder.AddExtensions("adp")
	'		cScanFolder.AddExtensions("mdb")
	'
	' ------------------------------------------------------------------
	Public Sub AddExtensions(sExtension)
		colExtensions.Add sExtension, True
	End Sub

	' ------------------------------------------------------------------
	'
	' Recursive scan : scan a folder and its subfolders and retrieve the list
	' of files having specific extensions. 
	'
	' Each file found will be outputted in a .txt file for later use
	'
	' ------------------------------------------------------------------
	Private Sub FileSearch(sFolderName)

		Dim objFSO, objFolder, objSubFolder, objFile, colFiles

		Set objFSO = CreateObject("Scripting.FileSystemObject")
		Set objFolder = objFSO.GetFolder(sFolderName) 

		Set colFiles = objFolder.Files

		For Each objFile in colFiles
			' Check if the file's extension is one that we're searching for
			If colExtensions.Exists(objfso.GetExtensionName(objFile)) Then
				' Yes => output the full filename in the output file
				objOutput.WriteLine(objFile.Path)
			End If
		Next

		' Call the subroutine for each subfolder within the original folder
		For Each objSubFolder In objFolder.SubFolders
			Call FileSearch(objSubFolder.Path)
		Next	

	End Sub

	' ------------------------------------------------------------------
	' Start the search, scan the "sSearchFolder" and subfolders
	' Example : 
	'
	'		Set cScanFolder = New clsScanFolder
	'		
	'		cScanFolder.AddExtensions("xlsx")
	'		cScanFolder.OutputFileName = "c:\temp\output.txt"
	'		cScanFolder.SearchFolder = "c:\temp"
	'
	'		Call cScanFolder.Run()
	'
	'		Set oShell = WScript.CreateObject ("WScript.Shell")
	'		oShell.run "notepad.exe c:\temp\output.txt"
	'		Set oShell = Nothing
	'
	' ------------------------------------------------------------------
	Public Function Run() 

		If bVerbose Then 
			wscript.echo vbCrLf & "=== clsScanFolder::Run ===" & vbCrLf
		End If

		Call FileSearch(sSearchFolder)
		Run =  true

	End Function

End Class