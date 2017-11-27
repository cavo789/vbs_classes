' ===========================================================================
'
' Author : Christophe Avonture
' Date   : November 2017
'
' MS Excel helper
'
' This class provide functionnalities like the function OpenCSV() 
' 
' ===========================================================================

Option Explicit

Class clsMSExcel

	Private oApplication
	Private bVerbose 

	Public Property Let verbose(bYesNo)
		bVerbose = bYesNo
	End Property

	Private Sub Class_Initialize()
		bVerbose = False

		Set oApplication = CreateObject("Excel.Application")

		oApplication.Visible = True
		oApplication.ScreenUpdating = True		
	End Sub

	Private Sub Class_Terminate()
		Set oApplication = Nothing
	End Sub

	Public Sub Quit()
		oApplication.Quit
	End Sub

	' Open a CSV file, correctly manage the split into columns,
	' add a title, rename the tab
	Public Sub OpenCSV(ByVal sFileName, ByVal sTitle, ByVal sSheetCaption)

		Dim objFSO
		Dim wCol

		Set objFSO = CreateObject("Scripting.FileSystemObject")

		If (objFSO.FileExists(sFileName)) Then

			If bVerbose Then
				wscript.echo "Open " & sFileName & " (clsMSExcel::OpenCSV)"
			End If

			' 1 =  xlDelimited
			' Delimiter is ";"
			oApplication.Workbooks.OpenText sFileName,,,1,,,,,,,True,";"

			' If a title has been specified, add quickly a small templating
			If (Trim(sTitle) <> "") Then

				With oApplication.ActiveSheet

					' Get the number of colunms in the file
					wCol = .Range("A1").CurrentRegion.Columns.Count

					.Range("1:3").insert
					.Range("A2").Value = Trim(sTitle)

					With .Range(.Cells(2, 1), .Cells(2, wCol))
						.HorizontalAlignment = 7 ' xlCenterAcrossSelection
						.font.bold = True
						.font.size = 14
					End with

					.Cells(4,1).AutoFilter
					
					.Columns.EntireColumn.AutoFit

					.Cells(5,1).Select

				End with

				oApplication.ActiveWindow.DisplayGridLines = False
				oApplication.ActiveWindow.FreezePanes = true
				
			End If

			If (Trim(sSheetCaption) <> "") Then 
				oApplication.ActiveSheet.Name = sSheetCaption
			End If

		End If

	End Sub

	Public Sub CloseFile
		oApplication.Workbooks(1).Close False
		
	End Sub

	Public Sub SaveFile(ByVal sFileName)
	
		If (bVerbose) Then
			wScript.Echo "Save file " & sFileName & " (clsMSExcel::SaveFile)"
		End If
		
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		If (objFSO.FileExists(sFileName)) Then
			Call objFSO.DeleteFile (sFileName)
		End If
		Set objFSO = Nothing

		oApplication.ActiveWorkbook.SaveAs sFileName, 51 ' xlWorkbookDefault
		
	End Sub

End Class