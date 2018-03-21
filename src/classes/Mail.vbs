' =======================================================
'
' Author : Christophe Avonture
' Date	: December 2017
'
' Provide a simple solution to work with emails
'
' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Mail.md
'
' =======================================================

Option Explicit

Class clsMail

	Private oApplication
	Private bVerbose
	Private sTo, sSubject, sHTMLBody
	Private bAppHasBeenStarted

	Public Property Let verbose(bYesNo)
		bVerbose = bYesNo
	End Property

	Public Property Let Recipient(ByVal sValue)
		sTo = sValue
	End Property

	Public Property Let Subject(ByVal sValue)
		sSubject = sValue
	End Property

	Public Property Let HTMLBody(ByVal sMessage)
		sHTMLBody = sMessage
	End Property

	Private Sub Class_Initialize()
		bVerbose = False
		Set oApplication = Nothing
		bAppHasBeenStarted = False
	End Sub

	Private Sub Class_Terminate()
		Set oApplication = Nothing
	End Sub

	' --------------------------------------------------------
	' Initialize the oApplication object variable : get a pointer
	' to the current Outlook.exe app if already in memory or start
	' a new instance.
	'
	' If a new instance has been started, initialize the variable
	' bAppHasBeenStarted to True so the rest of the script knows
	' that Outlook should then be closed by the script.
	' --------------------------------------------------------
	Private Sub Instantiate

		If (oApplication Is Nothing) Then

			On error Resume Next

			Set oApplication = GetObject(,"Outlook.Application")

			If (Err.number <> 0) or (oApplication Is Nothing) Then
				Set oApplication = CreateObject("Outlook.Application")
				' Remember that Outlook has been started by
				' this script ==> should be released
				bAppHasBeenStarted = True
			End if

			Err.clear

			On error Goto 0

		End if

	End Sub

	' -----------------------------------------------------
	' Prepare and display an email window.
	'
	' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Mail.md#display
	' -----------------------------------------------------
	Public Function Display()

		Dim objItem

		If (oApplication Is Nothing) Then
			Call Instantiate
		End if

		Set objItem = oApplication.CreateItem(0) ' MailItem

		With objItem
			.To = sTo
			.Subject = sSubject
			.ReadReceiptRequested = False
			.HTMLBody = sHTMLBody
		End With

		objItem.Display

		Set objItem = Nothing

		If bAppHasBeenStarted Then
			oApplication.Quit
			Set oApplication = Nothing
		End if

	End Function

End Class