' ========================================================
'
' Author : Christophe Avonture
' Date	: November / December 2017
'
' MS Excel helper
'
' This class provide functionnalities to facilitate automation of
' MS Excel
'
' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSExcel.md
'
' Changes
' =======
'
' March 2018 - Improve OpenCSV method
'
' ========================================================

Option Explicit

Class clsMSExcel

    Private oApplication
    Private sFileName
    Private bVerbose, bEnableEvents, bDisplayAlerts

    Private bAppHasBeenStarted

    Public Property Let verbose(bYesNo)
        bVerbose = bYesNo
    End Property

    Public Property Let EnableEvents(bYesNo)
        bEnableEvents = bYesNo

        If Not (oApplication Is Nothing) Then
            oApplication.EnableEvents = bYesNo
        End if
    End Property

    Public Property Let DisplayAlerts(bYesNo)
        bDisplayAlerts = bYesNo

        If Not (oApplication Is Nothing) Then
            oApplication.DisplayAlerts = bYesNo
        End if

    End Property

    Public Property Let FileName(ByVal sName)
        sFileName = sName
    End Property

    Public Property Get FileName
        FileName = sFileName
    End Property

    Public Property Let caption(ByVal sValue)
        If Not (oApplication Is Nothing) Then
            oApplication.Caption = sValue
        End If
    End Property

    ' Make oApplication accessible
    Public Property Get app
        Set app = oApplication
    End Property

    Private Sub Class_Initialize()
        bVerbose = False
        bAppHasBeenStarted = False
        bEnableEvents = False
        bDisplayAlerts = False
        Set oApplication = Nothing
    End Sub

    Private Sub Class_Terminate()
        Set oApplication = Nothing
    End Sub

    ' --------------------------------------------------------
    ' Initialize the oApplication object variable : get a pointer
    ' to the current Excel.exe app if already in memory or start
    ' a new instance.
    '
    ' If a new instance has been started, initialize the variable
    ' bAppHasBeenStarted to True so the rest of the script knows
    ' that Excel should then be closed by the script.
    ' --------------------------------------------------------
    Public Function Instantiate()

        If (oApplication Is Nothing) Then

            On error Resume Next

            Set oApplication = GetObject(,"Excel.Application")

            If (Err.number <> 0) or (oApplication Is Nothing) Then
                Set oApplication = CreateObject("Excel.Application")
				
				' Still nothing? Excel is thus not installed on the user's
				' computer ==> the user should installed Excel before
				If (oApplication Is Nothing) Then
					wScript.echo "Excel is required, please install Excel first (clsMSExcel::Instantiate)"
					wScript.Quit 
				End If
				
                ' Remember that Excel has been started by
                ' this script ==> should be released
                bAppHasBeenStarted = True
            End If

            oApplication.EnableEvents = bEnableEvents
            oApplication.DisplayAlerts = bDisplayAlerts

            Err.clear

            On error Goto 0

        End If

        ' Return True if the application was created right
        ' now
        Instantiate = bAppHasBeenStarted

    End Function

    ' --------------------------------------------------------
    ' Be sure Excel is visible
    ' --------------------------------------------------------
    Public Sub MakeVisible

        Dim objShell

        If Not (oApplication Is Nothing) Then

            With oApplication

                .Application.ScreenUpdating = True
                .Application.Visible = True
                .Application.DisplayFullScreen = False

                .WindowState = -4137 ' xlMaximized

            End With

            Set objShell = CreateObject("WScript.Shell")
            objShell.appActivate oApplication.Caption
            Set objShell = Nothing

        End If

    End Sub

    Public Sub Quit()
        If not (oApplication Is Nothing) Then
            oApplication.Quit
        End If
    End Sub

    ' --------------------------------------------------------
    ' Open a CSV file, correctly manage the split into columns,
    ' add a title, rename the tab
    '
    ' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSExcel.md#opencsv
    ' --------------------------------------------------------
    Public Sub OpenCSV(sTitle, sSheetCaption)

        Dim objFSO
        Dim wCol

        If bVerbose AND (sFileName = "") Then
            wScript.echo "Error, you need to initialize the " & _
                "filename first", " (clsMSExcel::OpenCSV)"
            Exit sub
        End If

        Set objFSO = CreateObject("Scripting.FileSystemObject")

        If (objFSO.FileExists(sFileName)) Then

            If bVerbose Then
                wScript.echo "Open " & sFileName & _
                    " (clsMSExcel::OpenCSV)"
            End If

            If (oApplication Is Nothing) Then
                Call Instantiate()
            End If

            ' 1 =  xlDelimited
            ' Delimiter is ";"
            oApplication.Workbooks.OpenText sFileName,,,1,,,,,,,True,";"

            ' If a title has been specified,
            ' add quickly a small templating
            If (Trim(sTitle) <> "") Then

                With oApplication.ActiveSheet

                    ' Get the number of colunms in the file
                    wCol = .Range("A1").CurrentRegion.Columns.Count

                    .Range("1:3").insert
                    .Range("A2").Value = Trim(sTitle)

                    With .Range(.Cells(2, 1), .Cells(2, wCol))
                        ' 7 = xlCenterAcrossSelection
                        .HorizontalAlignment = 7
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

        Set objFSO = Nothing

    End Sub

    ' --------------------------------------------------------
    ' Open a standard Excel file and allow to specify if the
    ' file should be opened in a read-only mode or not
    '
    ' 20190124 - Add sFileName as parameter and don't use the global variable
    ' --------------------------------------------------------
    Public Sub Open(sFileName, bReadOnly)

        If not (oApplication Is nothing) Then

            If bVerbose Then
                wScript.echo "Open " & sFileName & _
                    " (clsMSExcel::Open)"
            End If

            ' False = UpdateLinks
            oApplication.Workbooks.Open sFileName, False, _
                bReadOnly

        End If

    End sub

    ' --------------------------------------------------------
    ' Close the active workbook
    ' --------------------------------------------------------
    Public Sub CloseFile(sFileName)

        Dim wb
        Dim I
        Dim objFSO
        Dim sBaseName

        If Not (oApplication Is Nothing) Then

            Set objFSO = CreateObject("Scripting.FileSystemObject")

            If (sFileName = "") Then
                If Not (oApplication.ActiveWorkbook Is Nothing) Then
                    sFileName = oApplication.ActiveWorkbook.FullName
                End If
            End If

            If (sFileName <> "") Then

                If bVerbose Then
                    wScript.echo "Close " & sFileName & _
                        " (clsMSExcel::CloseFile)"
                End if

                ' Only the basename and not the full path
                sBaseName = objFSO.GetFileName(sFileName)

                On Error Resume Next
                Set wb = oApplication.Workbooks(sBaseName)
                If Not (err.number = 0) Then
                    ' Not found, workbook not loaded
                    Set wb = Nothing
                Else
                    If bVerbose Then
                        wScript.echo "	Closing " & sBaseName & _
                            " (clsMSExcel::CloseFile)"
                    End if
                    ' Close without saving
                    wb.Close False
                End if

                On Error Goto 0

            End If

            Set objFSO = Nothing

        End If

    End Sub

    ' --------------------------------------------------------
    ' Save the active workbook on disk
    '
    ' 20190124 - Add sFileName as parameter and don't use the global variable
    ' --------------------------------------------------------
    Public Sub SaveFile(sFileName)

        Dim wb, objFSO

        ' If Excel isn't loaded or has no active workbook, there
        ' is thus nothing to save.
        If Not (oApplication Is Nothing) Then

            Set objFSO = CreateObject("Scripting.FileSystemObject")

            On Error Resume Next

            Set wb = oApplication.Workbooks(objFSO.GetFileName(sFileName))

            If (Err.Number <> 0) Then
                Err.clear
                ' Perhaps the file isn't a .xlsx (.xlsm) file but an Addin$
                ' Try with the AddIns2 collection
                Set wb = oApplication.AddIns2(objFSO.GetFileName(sFileName))
            End If

            On Error Goto 0

            If Not (wb is Nothing) Then

                If (bVerbose) Then
                    wScript.echo "Save file " & sFileName & _
                        " (clsMSExcel::SaveFile)"
                End If

                If (wb.FullName = sFileName) Then
                    wb.Save
                Else
                    ' Don't specify extension because if we've opened
                    ' a .xlsm file and save the file elsewhere, we need
                    ' to keep the same extension
                    wb.SaveAs sFileName
                End If
            End IF

            Set wb = Nothing
            Set objFSO = Nothing

        End If

    End Sub

    ' --------------------------------------------------------
    ' Get the language of the MS Office interface
    ' Only return FR or NL; don't manage other languages
    ' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSExcel.md#getapplicationlanguage
    ' --------------------------------------------------------
    Public Function getApplicationLanguage(sDefault)

        Dim iLanguageID, bShouldClose
        Dim sValue

        ' Default value
        sValue = sDefault

        If (oApplication Is Nothing) Then
            bShouldClose = Instantiate()
        End If

        On error Resume Next

        ' 2 = msoLanguageIDUI
        iLanguageID = oApplication.LanguageSettings.LanguageID(2)

        If (Err.number <> 0) Then
            iLanguageID = 0
        End If

        Err.Clear

        On error Goto 0

        ' Quit Excel if it was started here, in this script
        If bShouldClose then
            oApplication.Quit
            Set oApplication = Nothing
        End If

        If (iLanguageID<>0) then
            ' Ok, we've found the language of MS Office

            If ((iLanguageID="1036") OR _
                (iLanguageID="2060") OR _
                (iLanguageID="3084") OR _
                (iLanguageID="4108") OR _
                (iLanguageID="5132")) Then
                ' MS Office has been installed in French
                sValue="FR"
            ElseIf (iLanguageID="1043") OR (iLanguageID="2067") Then
                ' MS Office has been installed in Dutch
                sValue="NL"
            End If

        End If

        getApplicationLanguage = sValue

    End Function

    ' --------------------------------------------------------
    ' Check if a specific file is already opened in Excel
    ' This function will return True if the file is already loaded.
    '
    ' 20190124 - Add sFileName as parameter and don't use the global variable
    ' --------------------------------------------------------
    Public Function IsLoaded(sFileName)

        Dim bLoaded, bShouldClose
        Dim bCheckAddins2
        Dim I, J

        bLoaded = false

        If (oApplication Is Nothing) Then
            bShouldClose = Instantiate()
        End If

        On Error Resume Next

        If bVerbose Then
            wScript.echo "Check if " & sFileName & _
                " is already loaded (clsMSExcel::IsLoaded)"
        End If

        If (Right(sFileName, 5) = ".xlam") Then

            ' The AddIns2 collection only exists since MSOffice
            ' 2014 (version 14)
            On Error Resume Next
            J = oApplication.AddIns2.Count
            bCheckAddins2 = (Err.Number = 0)
            On Error Goto 0

            If (bCheckAddins2) then

                J = oApplication.AddIns2.Count

                If J > 0 Then
                    For I = 1 To J
                        If (StrComp(oApplication.AddIns2(I).FullName,sFileName,1)=0) Then
                            bLoaded = True
                            Exit For
                        End If
                    Next ' For I = 1 To J
                End If

            End If ' If (oApplication.version >=14) then

        Else ' If (Right(sFileName, 5) = ".xlam") Then

            ' It's a .xls, .xlsm, ... file, not an AddIn
            J = oApplication.Workbooks.Count

            If J > 0 Then
                For I = 1 To J
                    If (StrComp(oApplication.Workbooks(I).FullName,sFileName,1)=0) Then
                        bLoaded = True
                        Exit For
                    End If
                Next ' For I = 1 To J
            End If ' If J > 0 Then

        End If ' If (Right(sFileName, 5) = ".xlam") Then

        On Error Goto 0

        ' Quit Excel if it was started here, in this script
        If bShouldClose then
            oApplication.Quit
            Set oApplication = Nothing
        End If

        IsLoaded = bLoaded

    End Function

    ' --------------------------------------------------------
    ' Display the list of references used by the file
    '
    ' bOnlyExternal 	Y/N - Restrict the list to only external
    '					references
    ' bOnlyXLAM			Y/N - Restrict the list to only references
    '					to *.xlam file
    '
    ' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSExcel.md#references_listall
    '
    ' 20190124 - Add sFileName as parameter and don't use the global variable
    ' --------------------------------------------------------
    Public Sub References_ListAll(sFileName, bOnlyExternal, bOnlyXLAM)

        Dim wb, ref
        Dim bShow, bEmpty
        Dim objFSO

        If Not (oApplication Is Nothing) Then

            Set objFSO = CreateObject("Scripting.FileSystemObject")

            Set wb = oApplication.Workbooks(objFSO.GetFileName(sFileName))
            bEmpty = True

            If Not (wb Is Nothing) Then
                wScript.echo "List of references in " & sFileName
                wScript.echo ""

                For Each ref In wb.VBProject.References

                    bShow = true

                    If (bOnlyExternal) Then
                        bShow = (ref.BuiltIn = False)
                    End If

                    If (bShow AND bOnlyXLAM) Then
                        bShow = (right(ref.FullPath,5) = ".xlam")
                    End If

                    if bShow then

                        bEmpty = False

                        wScript.echo " Name " & ref.Name
                        wScript.echo " Built In: " & ref.BuiltIn
                        wScript.echo " Full Path: " & ref.FullPath
                        wScript.echo " Is Broken: " & ref.IsBroken
                        wScript.echo " Version: " & ref.Major & "." & ref.Minor
                        wScript.echo ""

                    End If
                Next
            End If

            If bEmpty Then
                wScript.echo "The file didn't use references"
            End If

            Set wb = Nothing
            Set objFSO = nothing

        End If
    End sub

    ' --------------------------------------------------------
    ' Remove a specific reference. For instance, remove a linked
    ' .xlam library from the list of references used by the file
    '
    ' sAddin   The name of the reference; without the extension
    '          For instance "MyLibrary" (and not MyLibrary.xlam)
    '
    ' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSExcel.md#references_remove
    '
    ' Note: Make sure Events are disabled before calling this subroutine
    '       ==> .EnableEvents = False
    '
    ' 20190124 - Add sFileName as parameter and don't use the global variable
    ' --------------------------------------------------------
    Public Sub References_Remove(sFileName, sAddin)

        Dim wb, ref
        Dim objFSO
        Dim sRefFullName, sBaseName

        If Not (oApplication Is Nothing) Then

            Set objFSO = CreateObject("Scripting.FileSystemObject")

            ' sAddin should be a relative filename -
            ' Without the extension !
            If Instr(sAddin, "\")>0 Then
                sAddin = objFSO.GetBaseName(sAddin)
            End If

            If (StrComp(Right(sAddIn, 5), ".xlam", 1) = 0) Then
                sAddIn = Left(sAddIn, Len(sAddIn) - 5)
            End If

            'If bVerbose Then
            '    wScript.echo "Try to remove " & sAddin & " from the references"
            'End if

            ' IF STILL, EXCEL RUN A MACRO, IT'S MORE PROBABLY DUE TO A RIBBON
            ' PRESENT IN THE FILE AND AN "ONLOAD" SUBROUTINE.
            ' In this case, update the OnLoad code and check if
            ' Application.EnableEvents is equal to False and in this case,
            ' don't run your onLoad code; exit your subroutine.
            ' Add something like here below in the top of your subroutine:
            '
            '       If Not (Application.EnableEvents) Then
            '           Exit Sub
            '       End If
            '
            ' ALSO MAKE SURE TO NOT START EXCEL VISIBLE: THE RIBON IS LOADED
            ' IN THAT CASE

            Set wb = oApplication.Workbooks(objFSO.GetFileName(sFileName))

            If Not (wb Is Nothing) Then

                With wb

                    For Each ref In .VBProject.References

                        'If bVerbose Then
                        '    wScript.echo "   Found " & ref.Name & _
                        '        " (clsMSExcel::References_Remove)"
                        'End if

                        If (ref.Name = sAddIn) Then

                            If bVerbose Then
                                wScript.echo "      Remove the addin " & _
                                    "(clsMSExcel::References_Remove)"
                            End If

                            ' Get the fullpath of the reference
                            sRefFullName = ref.FullPath

                            .VBProject.References.Remove ref
                            .Save

                            ' --------------------------------------
                            ' Once unloaded, close the .xlam file
                            ' This should be made by closing the
                            ' filename (addin.xlam) and not just
                            ' the name (addin) or the fullname
                            ' So get the filename
                            sBaseName = objFSO.GetFileName(sRefFullName)

                            If bVerbose Then
                                wScript.echo "Unload " & sBaseName & _
                                " (clsMSExcel::References_Remove)"
                            End If

                            Call oApplication.Workbooks(sBaseName).Close

                            Exit For

                        End If

                    Next
                End With

            End If

            Set wb = Nothing
            Set objFSO = Nothing

        End If

    End Sub

    ' --------------------------------------------------------
    ' Add an addin to the list of references.
    ' For instance, add MyAddin.xlam to the MyInterface.xlsm
    '
    ' sAddin  The full filename to the addin to add as reference
    '
    ' 20190124 - Add sFileName as parameter and don't use the global variable
    ' --------------------------------------------------------
    Public Sub References_AddFromFile(sFileName, sAddinFile)

        Dim bReturn
        Dim wb, ref

        bReturn = true

        Dim objFSO

        If Not (oApplication Is Nothing) Then

            Set objFSO = CreateObject("Scripting.FileSystemObject")

            Set wb = oApplication.Workbooks(objFSO.GetFileName(sFileName))

            If Not (wb Is Nothing) Then

                If bVerbose Then
                    wScript.echo "Add a reference to " & sAddInFile
                End if

                wb.VBProject.References.AddFromFile sAddInFile

            End If

        End If

    End Sub

End Class
