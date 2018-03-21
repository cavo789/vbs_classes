' ==============================================================
'
' Author : Christophe Avonture
' Date	: December 2017
'
' Helpers to help working with folders
'
' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Folders.md
' ==============================================================

Option Explicit

Class clsFolders

	Dim objFSO, objFile

	Private bVerbose

	Public Property Let verbose(bYesNo)
		bVerbose = bYesNo
	End Property

	Private Sub Class_Initialize()
		bVerbose = False
		Set objFSO = CreateObject("Scripting.FileSystemObject")
	End Sub

	Private Sub Class_Terminate()
		Set objFSO = Nothing
	End Sub

	Public Function exists(sFolderName)
		exists = objFSO.FolderExists(sFolderName)
	End Function

	' --------------------------------------------------
	' Create a folder recursively
	' --------------------------------------------------
	Public Function create(sFolderName)

		create = false

		If Not exists(sFolderName) Then

			If create(objFSO.GetParentFolderName(sFolderName)) Then
				create = True

				If bVerbose Then
					wScript.echo "Create folder " & sFolderName & " " & _
						"(clsFolders::create)"
				End If

				Call objFSO.CreateFolder(sFolderName)
			End If

		Else
			create = True
		End If

	End Function

	' --------------------------------------------------
	' Get the list of files with the specified extension
	' and return a Dictionary object with, for each file,
	' the number of lines in the file
	'
	' Parameters :
	'
	' sFolder : the folder to scan
	' sExtension : the extension to search (f.i. "txt")
	'
	' Remark : if files are big, this function can take a while
	' so just be patient
	'
	' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Folders.md#countlines
	' --------------------------------------------------
	Public Function countLines(sFolder, sExtension)

		Dim objDict
		Dim objFile, objContent
		Dim wCountLines

		Set objDict = CreateObject("Scripting.Dictionary")

		' Loop any files
		For Each objFile In objFSO.GetFolder(sFolder).Files

			If (LCase(objFSO.GetExtensionName(objFile.Name)) = sExtension) Then

				Set objContent = objFSO.OpenTextFile(sFolder & objFile.Name, 1)

				objContent.ReadAll

				wCountLines = objContent.Line

				objdict.Add objFile.Name, wCountLines

			End if

		Next

		Set objContent = Nothing
		Set objFile = Nothing

		Set countLines = objdict

	End Function

	' --------------------------------------------------
	' Return the current folder i.e. the folder from where
	' the script has been started
	'
	' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Folders.md#getcurrentfolder
	' --------------------------------------------------
	Public Function getCurrentFolder()

		Dim sFolder

		Set objFile = objFSO.GetFile(Wscript.ScriptName)
		sFolder = objFSO.GetParentFolderName(objFile) & "\"
		Set objFile = Nothing

		getCurrentFolder = sFolder

	End Function

End Class
