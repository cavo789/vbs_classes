' ==============================================================
'
' Author : Christophe Avonture
' Date	: November 2017
'
' Helpers to help working with files
'
' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Files.md
'
' ==============================================================

Option Explicit

Class clsFiles

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

	' --------------------------------------------------
	' Create a text file
	' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Files.md#createtextfile
	' --------------------------------------------------
	Public Sub CreateTextFile(ByVal sFileName, ByVal sContent)

		If bVerbose Then
			wScript.echo "Create file " & sFileName & " " & _
				"(clsFiles::CreateTextFile)"
		End If

		Set objFile = objFSO.CreateTextFile(sFileName, 2, True)
		objFile.Write sContent
		objFile.Close
		Set objFile = Nothing

	End Sub

	Public Function FileExists(ByVal sFileName)
		FileExists = objFSO.FileExists(sFileName)
	End Function

	' --------------------------------------------------
	' Determine if a file is Readonly or not
	' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Files.md#isreadonly
	' --------------------------------------------------
	Public Function IsReadOnly(ByVal sFileName)

		IsReadOnly = False

		If objFso.FileExists(sFileName) Then
			If objFso.GetFile(sFileName).Attributes And 1 Then
				IsReadOnly = True
			End If
		End If

	End Function

	' --------------------------------------------------
	' Set the ReadOnly attribute for a file.
	' Assuming that the file exists.
	' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Files.md#setreadonly
	' --------------------------------------------------
	Public Sub SetReadOnly(sFileName)

		Dim objFile

		On Error Resume Next

		Set objFile = objFSO.GetFile(sFileName)

		' Not sure that the connected user can change
		' file's attributes. If he can't don't raise an error
		objFile.Attributes = objFile.Attributes OR 1

		If Err.Number <> 0 Then
			Err.Clear
		End if

		On Error Goto 0

		Set objFile = Nothing

	End Sub

	' --------------------------------------------------
	' Remove the ReadOnly attribute for a file.
	' Assuming that the file exists.
	' --------------------------------------------------
	Public Sub SetReadWrite(sFileName)

		Dim objFile

		On Error Resume Next

		Set objFile = objFSO.GetFile(sFileName)

		' Not sure that the connected user can change
		' file's attributes. If he can't don't raise an error
		objFile.Attributes = objFile.Attributes XOR 1

		If Err.Number <> 0 Then
			Err.Clear
		End if

		On Error Goto 0

		Set objFile = Nothing

	End Sub

	Public Function Copy(ByVal sSource, ByVal sTarget)

		Dim bReadOnly

		If (FileExists(sSource)) Then
			wScript.echo "Copy " & sSource & " to " & sTarget

			bReadOnly = IsReadOnly(sTarget)

			If bReadOnly Then
				' Remove the read-only attribute
				SetReadWrite(sTarget)
			End If

			objFSO.CopyFile sSource, sTarget

			If bReadOnly Then
				' And set it again
				SetReadOnly(sTarget)
			End If

		End if
	End Function

End Class
