' ==============================================================
'
' Author : Christophe Avonture
' Date	 : June 2018
'
' Helpers for working with .vbs scripts
'
' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/Helpers.md
' ==============================================================

Option Explicit

Class clsHelper

	' When the user double-clic on a .vbs file (from Windows explorer f.i.)
	' the running process will be WScript.exe while it's CScript.exe when
	' the .vbs is started from the command prompt.
	'
	' This subroutine will check if the script has been started with cscript
	' and if not, will run the script again with cscript and terminate the
	' "wscript" version. This is usefull when the script generate a lot of
	' wScript.echo statements, easier to read in a command prompt.
	'
	' How to use : just put these three lines at the top of your script
	'
	'  Set cHelper = New clsHelper
	'	 Call cHelper.ForceCScriptExecution()
	'	 Set cHelper = Nothing

	Sub ForceCScriptExecution()

		Dim sArguments, Arg, sCommand

		If Not LCase(Right(WScript.FullName, 12)) = "\cscript.exe" Then

			' Get command lines paramters'
			sArguments = ""
			For Each Arg In WScript.Arguments
				sArguments=sArguments & Chr(34) & Arg & Chr(34) & Space(1)
			Next

			sCommand = "cmd.exe cscript.exe //nologo " & Chr(34) & _
			WScript.ScriptFullName & Chr(34) & Space(1) & Chr(34) & sArguments & Chr(34)

			' 1 to activate the window
			' true to let the window opened
			Call CreateObject("Wscript.Shell").Run(sCommand, 1, true)

			' This version of the script (started with WScript) can be terminated
			wScript.quit

		End If

	End Sub

End Class
