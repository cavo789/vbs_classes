' ==============================================================
' Author : Christophe Avonture
' Date	: December 2017
'
' Helpers to help debugging a .vbs script
'
' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Debug.md
' ==============================================================

Option Explicit

Class clsDebug

	Dim objFSO, objFile
	Dim bEnable, bOnlyConsole, bConsoleMode

	' If set to False, nothing will be echoed
	Public Property Let Enable(bYesNo)
		bEnable = bYesNo
	End Property

	Public Property Get Enable
		enable = bEnable
	End Property

	' By default, set on True : only output debugging when
	' the script has been started from the console (i.e. with
	' cscript.exe; not when started with wscript.exe)
	Public Property Let onlyConsole(bYesNo)
		bOnlyConsole = bYesNo
	End Property

	Public Property Get onlyConsole
		onlyConsole = bOnlyConsole
	End Property

	Private Sub Class_Initialize()
		bEnable = False
		bOnlyConsole = True

		' True when the script used to start the .vbs is cscript.exe
		bConsoleMode = InStr(1, WScript.FullName, "CScript", _
			vbTextCompare)

	End Sub

	Private Sub Class_Terminate()
	End Sub

	' Output a text information on the console or by using a
	' popup dialog (when started with wscript.exe)
	Public Function echo(sLine)

		' Output only if enabled and started with cscript.exe;
		' not wscript.exe
		If (Enable()) Then
			If (Not (onlyConsole()) OR (bConsoleMode)) Then
				wScript.echo sLine
			End If
		End If

	End Function

End Class
