' ===========================================================================
'
' Author : Christophe Avonture
' Date   : November 2017
'
' Helpers to help working with files
'
' ===========================================================================

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

	' Create a text file
	Public Sub CreateTextFile(ByVal sFileName, ByVal sContent)

		If bVerbose Then 
			wscript.echo "Create file " & sFileName & " " & _
				"(clsFiles::CreateTextFile)"
		End If

		Set objFile = objFSO.CreateTextFile(sFileName, 2, True)
		objFile.Write sContent
		objFile.Close
		Set objFile = Nothing	
		
	End Sub

End Class
