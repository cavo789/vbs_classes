' ===========================================================================
'
' Author : Christophe Avonture
' Date	: December 2017
'
' Windows functionnalities
'
' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Windows.md
'
' ===========================================================================

Option Explicit

Class clsWindows

Private bVerbose

	Public Property Let verbose(bYesNo)
		bVerbose = bYesNo
	End Property

	Private Sub Class_Initialize()
		bVerbose = False
	End Sub

	Private Sub Class_Terminate()
	End Sub

	' Read from the registry the user's preferred languages and
	' return the first mentionned between dutch or french
	' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Windows.md#getpreferredlanguage
	Public Function getPreferredLanguage(ByVal sDefault)

		Dim WSHShell, value, sValue
		Dim wPosFR, wPosNL

		On Error Resume Next

		Set WSHShell = CreateObject("WScript.Shell")
		value = WSHShell.RegRead("HKCU\Software\Microsoft\Internet Explorer\International\AcceptLanguage")

		' Default language if the detection can't works
		sValue = sDefault

		If (Err.number = 0) Then

			' value will contains supported languages
			' like f.i. "nl-BE;en-GB;fr-FR". We then need to
			' use this priorisation to determine if the user's
			' preferences is first French or Dutch

			wPosNL=InStr(value,"nl-")
			wPosFR=InStr(value,"fr-")

			if (wPosNL<1) Then
				' Dutch not found => the user prefer FR content
				sValue = "FR"
			ElseIf (wPosFR<1) Then
				' French not found => the user prefer NL content
				sValue = "NL"
			Else
				' Both languages has been found;
				' Which one has been set at first ?
				If (wPosFR<wPosNL) Then
					sValue = "FR"
				Else
					sValue = "NL"
				End if
			End if

		End if

		getPreferredLanguage = sValue

	End Function

	Public Function getUserName()

		Dim sUserName
		Dim oNetwork

		Set oNetwork = CreateObject("WScript.Network")
		sUserName = oNetwork.UserName
		Set oNetwork = nothing

		getUserName = sUserName

	End Function

End Class
