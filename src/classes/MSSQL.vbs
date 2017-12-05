' ====================================================================
'
' Author : Christophe Avonture
' Date	: November 2017
'
' MS SQL Server helper
'
' This class provide functionnalities for working with a
' SQL Server database
'
' Documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSSQL.md
' ====================================================================

Option Explicit

Class clsMSSQL

	Private bVerbose

	Private p_DSN
	Private p_ServerName, p_DBName, p_UserName, p_Password
	Private p_Delim

	Public Property Let verbose(bYesNo)
		bVerbose = bYesNo
	End Property

	Public Property Let ServerName(ByVal sValue)
		p_ServerName = sValue
	End Property

	Public Property Get ServerName()
		ServerName = p_ServerName
	End Property

	Public Property Let DatabaseName(ByVal sValue)
		p_DBName = sValue
	End Property

	Public Property Get DatabaseName()
		DatabaseName = p_DBName
	End Property

	Public Property Let UserName(ByVal sValue)
		p_UserName = sValue
	End Property

	Public Property Get UserName()
		UserName = p_UserName
	End Property

	Public Property Let Password(ByVal sValue)
		p_Password = sValue
	End Property

	Public Property Get Password()
		Password = p_Password
	End Property

	Public Property Get DSN()

		Dim sDSN

		sDSN = Replace(p_DSN, "%server%", p_ServerName)
		sDSN = Replace(sDSN, "%database%", p_DBName)

		' If no username was supplied, consider a trusted connection
		If (p_UserName = "") Then
			sDSN = sDSN & "Trusted_Connection=True;"
		Else
			sDSN = sDSN & "User Id={" & p_UserName & "};" & _
				"Password={" & p_Password & "};"
		End If

		DSN = sDSN

	End Property

	' Define the delimiter to use when outputting to a text
	' file (can be ";" or "," or "| " or ...)
	Public Property Let Delimiter(ByVal sDelimiter)

		p_Delim = sDelimiter

	End Property

	Private Sub Class_Initialize()

		bVerbose = False
		p_Delim = ";"

		' Initialize default connection string
		p_DSN = "Driver={SQL Server};Server={%server%};" & _
				"Database={%database%};"
		p_ServerName = ""
		p_DBName = ""
		p_UserName = ""
		p_Password = ""

	End Sub

	Private Sub Class_Terminate()
	End Sub

	Private Function IIf(expr, truepart, falsepart)
		IIf = falsepart
		If expr Then IIf = truepart
	End Function

	' --------------------------------------------------------------
	'
	' Get the system folder
	'
	' Return, for instance, "C:\Windows\System32\"
	'
	' --------------------------------------------------------------
	Private Function GetSystemFolder()

		Const SystemFolder= 1
		Dim fso
		Dim SysFolder
		Dim SysFolderPath

		Set fso = wscript.CreateObject("Scripting.FileSystemObject")
		Set SysFolder = fso.GetSpecialFolder(SystemFolder)

		GetSystemFolder = SysFolder.Path

		Set SysFolder = Nothing
		Set fso = Nothing

	End function

	' --------------------------------------------------------------
	' Read an entire table and generate a string with the table
	' content. Will probably give errors on big table since the
	' result is a string
	'
	' Parameters :
	'
	' * sTableName 				: name of the table
	' * bIsMarkdown (boolean) 	: should the ouput be a Markdown
	'								string ?
	'
	' --------------------------------------------------------------
	Private Function p_GetTableContent(ByVal sTableName, _
		ByVal bIsMarkdown)

		Const adOpenStatic = 3
		Const adLockOptimistic = 3

		Dim objConnection, rs, fld
		Dim sSQL, sReturn, sLine, sSeparator

		If bVerbose Then
			wScript.echo vbCrLf & "=== clsMSSQL::GetTableContent ===" & vbCrLf
		End If

		sReturn = ""

		Set objConnection = CreateObject("ADODB.Connection")
		Set rs = CreateObject("ADODB.Recordset")

		objConnection.Open DSN()

		sSQL = "SELECT * FROM " & sTableName & ""

		If bVerbose Then
			wScript.echo sSQL & vbCrLf
		End If

		rs.Open sSQL, objConnection, adOpenStatic, adLockOptimistic

		rs.MoveFirst

		sLine = Iif(bIsMarkdown, p_Delim, "")
		sSeparator = p_Delim

		For Each fld In rs.Fields
			sLine = sLine & fld.Name & _
				Iif(bIsMarkdown, " ", "") & p_Delim
			sSeparator = sSeparator & "--- " & p_Delim
		Next

		If Not (bIsMarkdown) Then
			sLine = Left(sLine, Len(sLine) - 1)
		End if

		sReturn = sLine & vbCrLf

		If (bIsMarkdown) Then
			' Add a line with separator only for Markdown output
			sReturn = sReturn & sSeparator & vbCrLf
		End if

		Do While Not rs.eof

			sLine = Iif(bIsMarkdown, p_Delim, "")

			For Each fld In rs.Fields
				sLine = sLine & fld.Value & _
					Iif(bIsMarkdown, " ", "") & p_Delim
			Next

			If Not (bIsMarkdown) Then
				sLine = Left(sLine, Len(sLine) - 1)
			End if
			sReturn = sReturn & sLine & vbCrLf

			rs.MoveNext

		Loop

		rs.close

		Set rs=nothing
		Set objConnection=nothing

		p_GetTableContent = sReturn

	End Function

	' --------------------------------------------------------------
	' Read an entire table and generate a string with the table
	' content. This string can f.i. be a CSV delimited string.
	'
	' Notes :
	'
	' 1. You can choose the delimiter (can be ";", ",", "|" or
	'		something else); just initialize the Delimiter() property
	' 2. Will probably give errors on big table since the
	' result is a string
	'
	' Parameter :
	'
	' * sTableName	: name of the table
	'
	' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSAccess.md#gettablecontent
	' --------------------------------------------------------------
	Public Function GetTableContent(ByVal sTableName)
		GetTableContent = p_GetTableContent(sTableName, false)
	End Function

	' --------------------------------------------------------------
	' Read an entire table and generate a string with the table
	' content. Respect the markdown format.
	'
	' Notes :
	'
	' 1. The delimiter will be "|" since the output should be
	'		a markdown string
	' 2. Will probably give errors on big table since the
	' result is a string
	'
	' Parameter :
	'
	' * sTableName	: name of the table
	'
	' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSAccess.md#gettablecontentmarkdown
	' --------------------------------------------------------------
	Public Function GetTableContentMarkdown(ByVal sTableName)

		' Overwrite the delimiter, should be the pipe followed
		' by a space
		p_Delim = "| "

		' Tells the private function that we need to output a .md
		' string
		GetTableContentMarkdown = p_GetTableContent(sTableName, true)
	End Function

	' --------------------------------------------------------------
	' Create a User DSN to access a database through ODBC
	'
	' Parameter :
	' Array where only the first parameter which is mandatory
	' Elements will be :
	'
	' * Name		: (mandatory) Name for the DSN
	'					f.i. "DSN_myDB"
	' * Description	: (optional)  A description for the DSN
	'					f.i. "DSN to use for ..."
	' * DriverDLL	: (optional)  Filename for the driver
	'					f.i. "sqlncli11" (default value)
	' * DriverName	: (optional)  A friendly name for the driver
	'					f.i. "SQL Server" (default value)
	'
	' @link : Based on  https://www.experts-exchange.com/questions/24773510/Create-User-DSN-using-vbscript.html
	'
	' See documentation : https://github.com/cavo789/vbs_scripts/blob/master/src/classes/MSAccess.md#sql_create_dsn
	'
	' Sample code :
	'
	' Set cMSSQL = New clsMSSQL
	'
	' cMSSQL.Verbose = True
	'
	' cMSSQL.ServerName = "ServerName"
	' cMSSQL.DatabaseName = "DatabaseName"
	' cMSSQL.UserName = "UserName"
	'
	' wScript.echo cMSSQL.CreateUserDSN(array("dsn DB"))
	'
	' Set cMSSQL = Nothing
	'
	' --------------------------------------------------------------
	Public Function CreateUserDSN(arrParams)

		Dim sName, sDescription, sDriver, sDriverName
		Dim wShell

		' Retrieve the name to give to the new DSN
		sName = arrParams(0)

		sDescription = "DSN " & sName
		If (UBound(arrParams)>1) Then
			sDescription = arrParams(1)
		End if

		sDriver = "sqlncli11" ' Default value SQL 11 Native Client
		If (UBound(arrParams)>2) Then
			sDriver = arrParams(2)
		End if
		sDriver = GetSystemFolder() & "\" & sDriver & ".dll"

		sDriverName = "SQL Server" ' Default value
		If (UBound(arrParams)>3) Then
			sDriverName = arrParams(3)
		End if

		Set wShell = WScript.CreateObject("WScript.Shell")

		Dim RegEdPath
		RegEdPath= "HKEY_CURRENT_USER\SOFTWARE\ODBC\ODBC.INI\" & _
			sName & "\"

		wShell.RegWrite  RegEdPath  , ""

		If bVerbose Then
			wScript.echo vbCrLf & "=== clsMSSQL::CreateDSN ===" & vbCrLf
			wScript.echo vbCrLf & "Server = " & ServerName()
			wScript.echo vbCrLf & "DB = " & DatabaseName()
			wScript.echo vbCrLf & "Driver = " & sDriver
			wScript.echo vbCrLf & "User = " & UserName() & vbCrLf
		End If

		wShell.RegWrite  RegEdPath & "Database", DatabaseName()
		wShell.RegWrite  RegEdPath & "Description", sDescription
		wShell.RegWrite  RegEdPath & "Driver", sDriver
		wShell.RegWrite  RegEdPath & "LastUser", UserName()
		wShell.RegWrite  RegEdPath & "Server", ServerName()

		wShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\ODBC\" & _
			"ODBC.INI\ODBC Data Sources\" & sName, sDriverName

		Set wShell = Nothing

		CreateUserDSN = true

	End Function

End Class