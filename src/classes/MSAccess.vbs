' ===========================================================================
'
' Author : Christophe Avonture
' Date   : November 2017
'
' MS Access helper
'
' This class provide functionnalities like the function GetListOfTables() 
' to get the list of tables in a database.
' 
' ===========================================================================

Option Explicit

Class clsMSAccess

	Private oApplication
	Private bVerbose 

	Public Property Let verbose(bYesNo)
		bVerbose = bYesNo		
	End Property

	Private Sub Class_Initialize()	
		bVerbose = False		
		Set oApplication = CreateObject("Access.Application")
		oApplication.Visible = True
	End Sub

	Private Sub Class_Terminate()
		oApplication.Quit
		Set oApplication = Nothing
	End Sub

	Public Sub OpenDatabase(sFileName)
		If (Right(sFileName,4) = ".adp") Then
			oApplication.OpenAccessProject sFileName
		Else
			oApplication.OpenCurrentDatabase sFileName
		End If
	End Sub

	Public Sub CloseDatabase()
		oApplication.CloseCurrentDatabase
	End Sub

	' Return the table type in a more readable format
	Private Function GetTableType(ByVal wType) 

		If (wType = 1) Then 
			GetTableType = "Local"
		ElseIf (wType = 4) Then 
			GetTableType = "ODBC"
		ElseIf (wType = 6) Then 
			GetTableType = "Linked"
		Else
			GetTableType = "Unknown"
		End if

	End Function

	' Verify that databases mentionned in the arrDBNames are well 
	' present and accessible to the user. Return false otherwise
	Private Function CheckIfFilesExists(ByRef arrDBNames)

		Dim objFSO
		Dim bReturn 
		Dim i, iMin, iMax 

		iMin = LBound(arrDBNames)
		iMax = UBound(arrDBNames)

		bReturn = True

		Set objFSO = CreateObject("Scripting.FileSystemObject")    

		For i = iMin To iMax
			If Not (objFSO.FileExists(arrDBNames(I))) Then
				bReturn = False
				wScript.Echo "ERROR - clsMSAccess::CheckIfFilesExists - " & _
					"File " & arrDBNames(I) & " not found"
			End if
		Next

		CheckIfFilesExists = bReturn

	End function

	' --------------------------------------------------------------
	'
	' Create a folder structure; create parents folder if not found
	' 	CreateFolderStructure("c:\temp\a\b\c\d\e") will create the full structure
	'	in one call
	'
	' --------------------------------------------------------------
	Private Sub CreateFolderStructure(ByVal sFolderName)

		Dim arrPart, sBaseName, sDirName
		Dim objFSO

		Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
		
		If Not (objFSO.FolderExists(sFolderName)) Then 

			' Explode the folder name in parts 
			arrPart = split(sFolderName, "\")

			sDirName = ""

			For Each sBaseName In arrPart

				If sDirName <> "" Then 
					sDirName = sDirName & "\"
				End If

				sDirName = sDirName & sBaseName

				If (objFSO.FolderExists(sDirName) = False) Then 
					objFSO.CreateFolder(sDirName & "\")
				End if

			Next

		End if

	End Sub

	' FieldTypeName
	' by Allen Browne, allen@allenbrowne.com. Updated June 2006.
	' copied from http://allenbrowne.com/func-06.html
	' (No license information found at that URL.)
	Private Function GetFieldTypeName(FieldType, FieldAttributes)
	
	    Dim sReturn
	
	    Select Case CLng(FieldType) 
	        Case 1: sReturn = "Yes/No"            ' dbBoolean
	        Case 2: sReturn = "Byte"                 ' dbByte
	        Case 3: sReturn = "Integer"           ' dbInteger
	       
			Case 4                                     ' dbLong
		   	 If (FieldAttributes And 17) = 0 Then ' dbAutoIncrField
		   		 sReturn = "Long Integer"
		   	 Else
		   		 sReturn = "AutoNumber"
		   	 End If
		    Case 5: sReturn = "Currency"         ' dbCurrency
		    Case 6: sReturn = "Single"             ' dbSingle
		    Case 7: sReturn = "Double"             ' dbDouble
		    Case 8: sReturn = "Date/Time"            ' dbDate
		    Case 9: sReturn = "Binary"             ' dbBinary (no interface)
		    Case 10                                     'dbText
		   	 If (FieldAttributes And 1) = 0 Then 'dbFixedField
		   		 sReturn = "Text"
		   	 Else
		   		 sReturn = "Text (fixed width)"        '(no interface)
		   	 End If
		    Case 11: sReturn = "OLE Object"     'dbLongBinary
		    Case 12                                     'dbMemo
		   	 If (FieldAttributes And 32768) = 0 Then ' dbHyperlinkField
		   		 sReturn = "Memo"
		   	 Else
		   		 sReturn = "Hyperlink"
		   	 End If
		    Case 15: sReturn = "GUID"                 'dbGUID

		    'Attached tables only: cannot create these in JET.
		    Case 16: sReturn = "Big Integer"        'dbBigInt
		    Case 17: sReturn = "VarBinary"       'dbVarBinary
		    Case 18: sReturn = "Char"                 'dbChar
		    Case 19: sReturn = "Numeric"           'dbNumeric
		    Case 20: sReturn = "Decimal"           'dbDecimal
		    Case 21: sReturn = "Float"               'dbFloat
		    Case 22: sReturn = "Time"                 'dbTime
		    Case 23: sReturn = "Time Stamp"      'dbTimeStamp

		    'Constants for complex types don't work prior to Access 2007.
		    'Case 101&: sReturn = "Attachment"         'dbAttachment
		    'Case 102&: sReturn = "Complex Byte"       'dbComplexByte
		    'Case 103&: sReturn = "Complex Integer"    'dbComplexInteger
		    'Case 104&: sReturn = "Complex Long"       'dbComplexLong
		    'Case 105&: sReturn = "Complex Single"     'dbComplexSingle
		    'Case 106&: sReturn = "Complex Double"     'dbComplexDouble
		    'Case 107&: sReturn = "Complex GUID"       'dbComplexGUID
		    'Case 108&: sReturn = "Complex Decimal"    'dbComplexDecimal
		    'Case 109&: sReturn = "Complex Text"       'dbComplexText		   
		   
	        Case Else: sReturn = "Field type " & fld.Type & " unknown"
	    End Select
	 
	    GetFieldTypeName = sReturn
	End Function
	
	' --------------------------------------------------------------
	'
	' Scan one or severall MS Access databases and retrieve the list 
	' of tables in these DBs.
	'
	' @arrDBNames : array - Contains the list of databases to scan
	' @bOnlyForeign : True/False - Return only attached tables or all
	'
	' Example = 
	'
	'    arr(0) = "c:\temp\db1.accdb"
	'    arr(1) = "c:\temp\db2.accdb"
	'    arr(2) = "c:\temp\db3.accdb"
	'
	'    wScript.Echo GetListOfTables(arr, True)
	'
	' --------------------------------------------------------------	
	Public Function GetListOfTables(ByRef arrDBNames, bOnlyForeign) 

		Dim i, iMin, iMax 
		Dim sSQL, sReturn
		Dim rs

		If IsArray(arrDBNames) Then

			' Before starting, just verify that files exists 
			' If no, show an error message and stop
			If CheckIfFilesExists(arrDBNames) Then
				' Ok, database(s) are well present, we can start

				sReturn = "Database Filename;Table name;Linked Database;" & _
					"Linked Table name;Table Type" & vbCrLf

				For i = iMin To iMax

					If bVerbose Then 
						wscript.echo "Process " & arrDBNames(I)
					End If

					Call OpenDatabase(arrDBNames(I))

					If bOnlyForeign then
						' Get only attached tables
						sSQL = "SELECT [Name] AS [TableName], Database, " & _
							"ForeignName " & _ 
							"FROM MsysObjects " & _
							"WHERE ForeignName IS NOT NULL " & _
							"ORDER BY Database, [Name];"
					Else
						' Get all tables : local or linked but not system ones
						'
						' MsysObjects.Type =
						' 	1 = Tables (Local)
						'	4 = Tables (Linked using ODBC)
						'	6 = Tables (Linked)

						sSQL = "SELECT [Name] AS [TableName], Database, " & _
							"ForeignName, Type, Flags " & _ 
							"FROM MsysObjects " & _
							"WHERE (MsysObjects.Name Not Like '~*') AND " & _
								"(MsysObjects.Name Not Like 'MSys*') AND " & _
								"(MsysObjects.Type IN (1, 4, 6)) AND " & _
								"(Flags Not In (-2146828288,-2147287040)) " & _
							"ORDER BY MsysObjects.Name"

					End if

					Set rs = oApplication.CurrentDB.OpenRecordset(sSQL, 4) 

					If rs.RecordCount <> 0 Then
						Do While Not rs.EOF
							sReturn = sReturn & _
								arrDBNames(I) & ";" & _
								rs.fields("TableName").Value & ";" & _
								rs.fields("Database").Value & ";" & _
								rs.Fields("ForeignName").Value & ";" & _
								GetTableType(rs.Fields("Type").Value) & _
								vbCrLf

							rs.MoveNext
							Loop
					End If

					If Not rs Is Nothing Then
						rs.Close
						Set rs = Nothing
					End If

					Call CloseDatabase

				Next

			End IF

		Else
			wScript.Echo "ERROR - clsMSAccess::GetListOfTables - " & _
				"You must provide an array with filenames."
		End If

		GetListOfTables = sReturn

	End Function
	
	' --------------------------------------------------------------
	'
	' Scan one or severall MS Access databases, retrieve the list 
	' of tables in these DBs and get the list of fields plus some properties
	' like the size and, for text fields, the shortest size and the longest one.
	'
	' @arrDBNames : array - Contains the list of databases to scan
	'
	' Example = 
	'
	'    arr(0) = "c:\temp\db1.accdb"
	'    arr(1) = "c:\temp\db2.accdb"
	'    arr(2) = "c:\temp\db3.accdb"
	'
	'    wScript.Echo GetFieldsList(arr)
	'
	' --------------------------------------------------------------	
	Public Function GetFieldsList(ByRef arrDBNames) 

		Dim i, iMin, iMax, sShortest, sLargest
		Dim sSQL, sReturn, sTableName, sType
		Dim objTable, objField, rs

		If IsArray(arrDBNames) Then

			' Before starting, just verify that files exists 
			' If no, show an error message and stop
			If CheckIfFilesExists(arrDBNames) Then
				' Ok, database(s) are well present, we can start

				sReturn = "Database Filename;Table name;Field name;Field type; " & _
					"Field size;Shortest size;Longest size" & vbCrLf

				For i = iMin To iMax

					If bVerbose Then 
						wscript.echo "Process " & arrDBNames(I)
					End If

					Call OpenDatabase(arrDBNames(I))

					oApplication.CurrentDB.TableDefs.Refresh

					For each objTable In oApplication.CurrentDB.TableDefs

						sTableName = objTable.Name

						' Ignore system and temporary tables
						If (lcase(Left(sTableName, 4))<>"msys") And (Left(sTableName, 1) <> "~") Then

							For Each objField In objTable.Fields   
							
								sShortest = ""
								sLargest = ""
								
								sType = GetFieldTypeName(objField.Type, objField.Attributes) 
								
								If (sType = "Text") Or (sType = "Memo") Then 
								
									sSQL = "SELECT " & _
										"Min(Len([" & objField.Name & "])) As Min, " & _
										"Max(Len([" & objField.Name & "])) As Max " & _
										"FROM [" & sTableName & "]"
									
									Set rs = oApplication.CurrentDB.OpenRecordset(sSQL, 4) 

									sShortest = rs.Fields("Min").Value 
									sLargest = rs.Fields("Max").Value 
									
				 					rs.Close
									
				 					Set rs = Nothing
									
								End If

								sReturn = sReturn & _
									arrDBNames(I) & ";" & _
									sTableName & ";" & _
									objField.Name & ";" & _
									sType & ";" & _
									objField.Size & ";" & _
									sShortest & ";" & sLargest & vbCrLf'
	    					Next

						End if

					Next

					Call CloseDatabase

				Next

			End IF

		Else
			wScript.Echo "ERROR - clsMSAccess::GetFieldsList - " & _
				"You must provide an array with filenames."
		End If

		GetFieldsList = sReturn

	End Function
		
	' --------------------------------------------------------------
	'
	' Scan one or severall MS Access databases and if table's name 
	' start with a given prefix (like "dbo_"), remove that prefix
	'
	' @arrDBNames : array - Contains the list of databases to scan
	' @sPrefix    : the prefix to remove (f.i. "dbo_")
	'
	' Example = 
	'
	'    arr(0) = "c:\temp\db1.accdb"
	'    arr(1) = "c:\temp\db2.accdb"
	'    arr(2) = "c:\temp\db3.accdb"
	'
	'    wScript.Echo RemovePrefix(arr, "dbo_")
	'
	' --------------------------------------------------------------	
	Public Sub RemovePrefix(ByRef arrDBNames, sPrefix) 

		Dim i, iMin, iMax 
		Dim objTable
		Dim sNewName

		If IsArray(arrDBNames) Then

			' Before starting, just verify that files exists 
			' If no, show an error message and stop
			If CheckIfFilesExists(arrDBNames) Then
				' Ok, database(s) are well present, we can start

				For i = iMin To iMax

					If bVerbose Then 
						wscript.echo "Process database " & arrDBNames(I)
					End If

					Call OpenDatabase(arrDBNames(I))

					For Each objTable in oApplication.CurrentData.AllTables
						If (Left(objTable.Name, Len(sPrefix)) = sPrefix) then

							sNewName = Mid(objTable.Name, Len(sPrefix) + 1)

							If bVerbose Then
								wScript.Echo "   Rename [" & objTable.Name & "] " & _
									"to [" & sNewName & "]"
							End If	

							' 0 = acTable
							oApplication.DoCmd.Rename sNewName, 0, objTable.Name

						End If
					Next

					Call CloseDatabase

				Next

			End IF

		Else
			wScript.Echo "ERROR - clsMSAccess::RemovePrefix - " & _
				"You must provide an array with filenames."
		End If

	End Sub
	
	' --------------------------------------------------------------
	'
	' Open a database and export every forms, macros, modules and reports code
	'
	' @arrDBNames : array - Contains the list of databases to scan
	'
	' Example = 
	'
	'    arr(0) = "c:\temp\db1.accdb"
	'    arr(1) = "c:\temp\db2.accdb"
	'    arr(2) = "c:\temp\db3.accdb"
	'
	'    wScript.Echo Decompose(arr)
	'
	' --------------------------------------------------------------	
	Public Sub Decompose(ByRef arrDBNames, sExportPath)
	
		Dim i, iMin, iMax 
		Dim objFSO, obj
		Dim myComponent
		Dim sModuleType
		Dim sTempName, sOutFileName
		Dim sDBExtension, sDBName, sDBParentFolder

		If IsArray(arrDBNames) Then

			' Before starting, just verify that files exists 
			' If no, show an error message and stop
			If CheckIfFilesExists(arrDBNames) Then
				' Ok, database(s) are well present, we can start

				For i = iMin To iMax

					If bVerbose Then 
						wscript.echo "Process database " & arrDBNames(I)
					End If

					Call OpenDatabase(arrDBNames(I))

					Set objFSO = CreateObject("Scripting.FileSystemObject")

					sDBExtension = objFSO.GetExtensionName(arrDBNames(I))
					sDBName = objFSO.GetBaseName(arrDBNames(I))
					sDBParentFolder = objFSO.GetParentFolderName(arrDBNames(I))

					If (sExportPath = "") then
						sExportPath = sDBParentFolder & "\" & sDBName & "_" & _
							sDBExtension & "\src\"
					End If

					Call CreateFolderStructure(sExportPath)
					Call CreateFolderStructure(sExportPath & "Forms\")
					Call CreateFolderStructure(sExportPath & "Macros\")
					Call CreateFolderStructure(sExportPath & "Modules\")
					Call CreateFolderStructure(sExportPath & "Reports\")

					For Each obj In oApplication.CurrentProject.AllForms

						sOutFileName = obj.FullName & ".txt"
						sOutFileName = sExportpath & "Forms\" & sOutFileName

						If bVerbose Then 
							WScript.Echo "  Export form " & obj.FullName & " to " & sOutFileName
						End If

						' 2 = acForm
						oApplication.SaveAsText 2, obj.FullName, sOutFileName
						oApplication.DoCmd.Close 2, obj.FullName

					Next

					For Each obj In oApplication.CurrentProject.AllModules

						sOutFileName = obj.FullName & ".txt"
						sOutFileName = sExportpath & "Modules\" & sOutFileName

						If bVerbose Then 
							WScript.Echo "  Export module " & obj.FullName & " to " & sOutFileName
						End If

						' 5 = acModule
						oApplication.SaveAsText 5, obj.FullName, sOutFileName

    				Next

					For Each obj In oApplication.CurrentProject.AllMacros

						sOutFileName = obj.FullName & ".txt"
						sOutFileName = sExportpath & "Macros\" & sOutFileName

						If bVerbose Then 
							WScript.Echo "  Export macro " & obj.FullName & " to " & sOutFileName
						End If

						' 4 = acMacro
						oApplication.SaveAsText 4, obj.FullName, sOutFileName

					Next
					
				    For Each obj In oApplication.CurrentProject.AllReports

						sOutFileName = obj.FullName & ".txt"
						sOutFileName = sExportpath & "Reports\" & sOutFileName

						If bVerbose Then 
							WScript.Echo "  Export report " & obj.FullName & " to " & sOutFileName
						End If

						' 3 = acReport
						oApplication.SaveAsText 3, obj.FullName, sOutFileName

					Next

					Call CloseDatabase

				Next 

			End If

		Else
			wScript.Echo "ERROR - clsMSAccess::Decompose - " & _
				"You must provide an array with filenames."
		End If

	End Sub 
	
End Class