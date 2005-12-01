Option Strict Off
Option Explicit On 
Imports Microsoft.VisualBasic
Imports Microsoft.VisualBasic.Compatibility
Imports System.Data.Odbc
Imports System.Drawing

Public Class CMitaShared
	Dim mitaData As pscMitaData.CMitaData
	Dim mitaConnect As pscMitaConnect.CMitaConnect
	Dim mitaSystem As pscMitaSapSystem.CMitaSapSystem
	Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	Private mvarRestore As Boolean = True
	Private mvarErrorDescription As String = ""

	Public WriteOnly Property center() As Boolean
		Set(ByVal Value As Boolean)
			mvarRestore = Not Value
		End Set
	End Property

	Public Function iniGetSections(ByRef iniFile As String) As String()
		Dim f1 As Integer
		Dim s() As String
		Dim sCnt As Integer
		Dim a As String
		f1 = FreeFile()
		If Dir$(iniFile) = "" Then Return Nothing
		FileOpen(f1, iniFile, OpenMode.Input)
		sCnt = -1
		Do
			If EOF(f1) Then Exit Do
			a = LineInput(f1)
			If Left(a, 1) = "[" Then
				sCnt = sCnt + 1
				ReDim Preserve s(sCnt)
				s(sCnt) = a
			End If
		Loop
		FileClose(f1)
		iniGetSections = VB6.CopyArray(s)
	End Function

	Public Function iniReadSection(ByRef section As String, ByRef iniFile As String, ByRef concat As Boolean) As String()
		Dim l As Integer
		Dim a As String
		Dim x As Integer
		Dim tmp() As String
		Dim keys() As String
		Dim spl() As String
		Dim keyCount As Integer
		Dim content As String
		Dim actKey As String
		Dim wrkKey As String
		Dim ind As Double
		l = 1024
		Do
			a = Space(l)
			x = GetPrivateProfileSection(section, a, l, iniFile)
			If x = 0 Then
				If l = 1024 Then Return Nothing
				Exit Do
			End If
			If x < l - 2 Then Exit Do
			l = l + 1024
		Loop
		a = Left(a, x - 1)
		tmp = Split(a, Chr(0))
		actKey = ""
		wrkKey = ""
		keyCount = -1
		content = ""
		If concat Then
			For l = 0 To UBound(tmp)
				spl = Split(tmp(l), "=", 2)
				wrkKey = Trim(spl(0))
				If wrkKey <> "DSN" Then
					x = InStr(wrkKey, "_")
					ind = Val(Mid(wrkKey, x + 1))
					If ind = 1.0 Then
						If content <> "" Then
							content = content
							keyCount = keyCount + 1
							ReDim Preserve keys(keyCount)
							keys(keyCount) = actKey & "=" & Replace(content, "'", "~")
						End If
						content = LTrim(spl(1))
						actKey = Left(wrkKey, x - 1)
					Else
						If wrkKey.StartsWith("@") Then
							actKey = wrkKey
						End If
						content = content & spl(1)
					End If
				End If
			Next l
			keyCount = keyCount + 1
			ReDim Preserve keys(keyCount)
			keys(keyCount) = actKey & "=" & Replace(content, "'", "~")
			iniReadSection = VB6.CopyArray(keys)
		Else
			iniReadSection = VB6.CopyArray(tmp)
		End If
	End Function


	Public Function iniReadSingle(ByRef AppName As String, ByRef KeyName As String, ByRef defaultValue As String, ByRef iniFileName As String) As String
		Dim buffer As String
		Dim a As Integer
		Dim IniFil As String
		IniFil = iniFileName
		Dim l As Integer
		If IniFil = "" Then
			IniFil = Replace(System.Reflection.Assembly.GetExecutingAssembly.Location, ".exe", ".ini", , , CompareMethod.Text)
			If Dir(IniFil) = "" Then
				IniFil = System.Reflection.Assembly.GetExecutingAssembly.GetName.Name & ".ini"
			End If
		End If
		iniReadSingle = ""
		l = 256
		Do
			buffer = Space(l)
			a = GetPrivateProfileString(AppName, KeyName, defaultValue, buffer, l, IniFil)
			If a = 0 Then Exit Function
			If a < l - 2 Then Exit Do
			l = l + 256
		Loop
		buffer = Left(buffer, a)
		iniReadSingle = Trim(buffer)
	End Function

	Public Shared Function DoXor(ByRef inp As String, Optional ByRef x As String = Nothing) As String
		Dim a As String
		Dim b As String
		Dim z As String
		Dim i As Integer
		Dim C As Byte
		Dim src As String
		Dim isScrambled As Boolean
		If inp = "" Then Return ""
		If Left(inp, 2) = Chr(255) & Chr(255) Then
			src = Mid$(inp, 3)
			isScrambled = True
			a = ""
		Else
			src = inp
			isScrambled = False
			a = Chr(255) & Chr(255)
		End If
		If IsNothing(x) Then
			z = "RJHATMORJHATPSRJHATPAKRJHATBLZRJHATISZRJHATPAPRJHATPLZARJHATPLZRJHATSTATRJHATTXTRJHATPSIRJHATERR"
		Else
			z = x
		End If
		b = ""
		While Len(b) < Len(src) : b = b & z : End While
		z = b

		For i = 1 To Len(src)
			C = CByte(Asc(Mid(src, i, 1)))
			If isScrambled Then
				Select Case C
					Case 254 : C = 13
					Case 255 : C = 0
					Case 253 : C = 10
					Case 252 : C = 34
					Case 251 : C = 9
					Case 250 : C = 39
					Case Else
				End Select
				C = C Xor CByte(Asc(Mid(z, i, 1)))
			Else
				C = C Xor CByte(Asc(Mid(z, i, 1)))
				Select Case C
					Case 10 : C = 253
					Case 34 : C = 252
					Case 13 : C = 254
					Case 0 : C = 255
					Case 9 : C = 251
					Case 39 : C = 250
					Case Else
				End Select
			End If
			'    If C < 40 Then
			'      a$ = a$ & "|"
			'      C = C + 40
			'    End If
			a = a & Chr(C)
			'    Debug.Print a$
		Next i
		DoXor = a
	End Function
	'Public Function buildHash(ByVal nameInput As String) As Integer
	'	Dim i As Integer
	'	Dim res As Integer = 0
	'	For i = 1 To Len(nameInput)
	'		res = res + Asc(Mid$(nameInput, i, 1))
	'	Next
	'	Return res Mod 256
	'End Function
	Public Sub buildVersionInfo(ByRef frm As System.Windows.Forms.Form)
		Dim i As Integer
		If InStr(frm.Text, "-") = 0 Then
			frm.Text = getExeName() & " - " & frm.Text
		End If
		i = InStr(frm.Text, "(")
		If i > 0 Then
			frm.Text = Trim$(Left$(frm.Text, i - 1))
		End If
		If mitaSystem.sapSystemNAME <> "" Then
			If mitaSystem.sapSystemVERSIONNAME <> "" Then
				frm.Text = frm.Text & " (" & mitaSystem.sapSystemNAME & ", " & mitaSystem.sapSystemVERSIONNAME & ", " & mitaSystem.rfcType & ")"
			Else
				frm.Text = frm.Text & " (" & mitaSystem.sapSystemNAME & ")"
			End If
		End If
	End Sub
	Public Sub RestPos(ByRef frm As System.Windows.Forms.Form, Optional ByVal versionInfo As Boolean = False, Optional ByRef size As Boolean = False)
		If mvarRestore Then
			Dim x As Double
			x = Val(GetSetting(mitaData.registryApplication, frm.Name, "Top", CStr(-1)))
			If x <> -1 Then
				frm.Top = x
				x = Val(GetSetting(mitaData.registryApplication, frm.Name, "Left", CStr(-1)))
				If x <> -1 Then frm.Left = x
				If Not IsNothing(size) Then
					If CBool(size) Then
						x = Val(GetSetting(mitaData.registryApplication, frm.Name, "Width", CStr(-1)))
						If x <> -1 Then frm.Width = x
						x = Val(GetSetting(mitaData.registryApplication, frm.Name, "Height", CStr(-1)))
						If x <> -1 Then frm.Height = x
					End If
				End If
			End If
		Else
			centerMe(frm)
		End If
		If versionInfo Then buildVersionInfo(frm)
		frm.Refresh()
		frm.Visible = True
		System.Windows.Forms.Application.DoEvents()
	End Sub
	Private Sub centerMe(ByVal frm As System.Windows.Forms.Form)
		Dim scL As System.Drawing.Rectangle = New Rectangle
		Dim scr As System.Windows.Forms.Screen = System.Windows.Forms.Screen.PrimaryScreen
		scL = scr.Bounds
		Dim x As Integer = (scL.Width - frm.Width) / 2
		Dim y As Integer = (scL.Height - frm.Height) / 2
		frm.Left = x
		frm.Top = y
	End Sub
	Public Property appName() As String
		Get
			Return mitaData.registryApplication
		End Get
		Set(ByVal Value As String)
			If IsNothing(mitaData) Then mitaData = New pscMitaData.CMitaData
			mitaData.registryApplication = Value
		End Set
	End Property
	Public Sub SavePos(ByRef frm As System.Windows.Forms.Form)
		If Not mvarRestore Then Exit Sub
		SaveSetting(mitaData.registryApplication, frm.Name, "Top", CStr(frm.Top))
		SaveSetting(mitaData.registryApplication, frm.Name, "Left", CStr(frm.Left))
		SaveSetting(mitaData.registryApplication, frm.Name, "Width", CStr(frm.Width))
		SaveSetting(mitaData.registryApplication, frm.Name, "Height", CStr(frm.Height))
	End Sub
	Public Function getExeName() As String
		Return System.Reflection.Assembly.GetEntryAssembly.GetName.Name
	End Function
	Public Function getSapIDFromName(ByVal connection As OdbcConnection, ByVal sapname As String) As Integer
		Dim sys() As String
		Dim id() As Integer
		Dim cnt As Integer
		Dim i As Integer
		If readDBSap(connection, sys, id, cnt) Then
			For i = 0 To cnt
				If sys(i) = sapname Then Return id(i)
			Next
		End If
		Return 0
	End Function
	Public Function readDBSap(ByVal odbc_connection As OdbcConnection, ByRef cSap() As String, ByRef cSapIndex() As Integer, ByRef cCount As Integer) As Boolean
		Dim query As String
		Dim a As String
		readDBSap = False
		cCount = -1
		Dim idbc As OdbcCommand = odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		query = "SELECT sapsystemid, sapname FROM " & mitaSystem.tableSapSystems
		query = query & " WHERE activ = 'Y'"
		query = query & " ORDER BY sapsystemid;"
		idbc.CommandText = query
		Try
			odbc_connection.Open()
			reader = idbc.ExecuteReader()
			While reader.Read
				cCount = cCount + 1
				ReDim Preserve cSap(cCount)
				ReDim Preserve cSapIndex(cCount)
				cSapIndex(cCount) = reader.GetDouble(0)
				cSap(cCount) = reader.GetString(1)
				readDBSap = True
			End While
		Catch
			mitaData.errorDescription = Err.Description & vbCrLf & query
			MsgBox(mitaData.errorDescription)
		End Try
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		odbc_connection.Close()
	End Function
	Public Function getBool(ByVal inp() As String) As Boolean
		If UBound(inp) = 0 Then Return True
		Return Not (StrComp(inp(1), "Off", CompareMethod.Text) = 0)
	End Function
	Private Function addID(ByVal host As String) As Integer
		Dim hostList() As String
		Dim hostCount As Integer
		Dim index As Integer = 0
		Dim query As String
		Dim result As Boolean
		Dim res As Boolean = getHostList(hostList, hostCount)
		Dim i As Integer
		For i = 0 To hostCount
			If host = hostList(i) Then
				query = "SELECT hostindex FROM " & mitaSystem.tableOnline
				query = query & " WHERE loginhost = '" & hostList(i) & "'"
				result = mitaConnect.queryNumber(query, mitaData.newID)
				Return mitaData.newID
			End If
		Next
		mitaData.newID = 0
		query = "SELECT hostindex FROM " & mitaSystem.tableOnline
		Dim exists As Boolean
		result = mitaConnect.queryExist(query, exists)
		If exists Then
			query = "SELECT MAX(hostindex) FROM " & mitaSystem.tableOnline
			result = mitaConnect.queryNumber(query, mitaData.newID)
		End If
		mitaData.newID = mitaData.newID + 100
		Return mitaData.newID
	End Function
	Public Function getHostList(ByRef cHosts() As String, ByRef cCount As Integer) As Boolean
		Dim query As String
		Dim sb As New System.Text.StringBuilder(1000)
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		getHostList = False
		cCount = -1
		sb.Append("SELECT DISTINCT loginhost from " & mitaSystem.tableOnline)
		sb.Append(" ORDER BY loginhost ASC")
		query = sb.ToString()
		idbc.CommandText = query
		On Error GoTo isErr
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			cCount = cCount + 1
			ReDim Preserve cHosts(cCount)
			cHosts(cCount) = CStr(reader.Item("loginhost"))
		End While
		getHostList = True
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
		Resume Exx
	End Function
	Public Function generateID(ByVal id As Integer) As Integer
		mitaData.tmpID = id
		mitaData.createID = mitaData.tmpID + addID(mitaData.createHost)
		Return mitaData.createID
	End Function
	Property systemSet() As pscMitaSapSystem.CMitaSapSystem
		Get
			Return mitaSystem
		End Get
		Set(ByVal Value As pscMitaSapSystem.CMitaSapSystem)
			mitaSystem = Value
		End Set
	End Property
	Property dataSet() As pscMitaData.CMitaData
		Get
			Return mitaData
		End Get
		Set(ByVal Value As pscMitaData.CMitaData)
			mitaData = Value
		End Set
	End Property
	Property connectSet() As pscMitaConnect.CMitaConnect
		Get
			Return mitaConnect
		End Get
		Set(ByVal Value As pscMitaConnect.CMitaConnect)
			mitaConnect = Value
		End Set
	End Property
	Public Function readIniSapSystemFromName(ByVal name As String, ByVal iniFile As String) As Boolean
		Dim i As Integer
		Dim sapKeys() As String = iniReadSection(name, iniFile, False)
		Dim x() As String
		For i = 0 To UBound(sapKeys)
			x = sapKeys(i).Split("="c)
			decodeKey(x(0).Trim, x(1).Trim)
		Next
	End Function
	Public Function setSapSystemVERSIONNAME() As Boolean
		mitaSystem.sapSystemVERSIONNAME = ""
		mvarErrorDescription = ""
		Dim result As Boolean
		Dim target As String
		Dim query As String = "SELECT sapversionname FROM pscsapversions"
		query = query & " WHERE sapversionid = " & mitaSystem.sapSystemVERSIONID & " AND activ = 'Y'"
		result = mitaConnect.queryString(query, target)
		If result Then
			mitaSystem.sapSystemVERSIONNAME = target
			Return True
		Else
			mvarErrorDescription = mitaConnect.errorDescription & vbCrLf & query
			Return False
		End If
	End Function
	Private Function readDBSap(ByRef query As String, Optional ByVal connString As String = Nothing) As Boolean
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		Dim sb As New System.Text.StringBuilder(1000)
		Dim i As Integer
		Dim nam As String
		Dim tstObject As Object
		readDBSap = False
		mvarErrorDescription = ""
		On Error GoTo isErr
		idbc.CommandText = query
		If Not IsNothing(connString) Then mitaConnect.dbConnectString = connString
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			For i = 0 To reader.FieldCount - 1
				nam = reader.GetName(i)
				tstObject = reader.Item(nam)
				decodeKey(nam, tstObject)
			Next i
			readDBSap = True
		End While
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mvarErrorDescription = Err.Description & vbCrLf & query
		Resume Exx
	End Function
	Private Sub decodeKey(ByVal nam As String, ByVal tstObject As Object)
		Select Case UCase(nam)
			Case "SAPSYSTEMID", "ID"
				mitaSystem.sapSystemId = CInt(tstObject)
			Case "SAPNAME"
				mitaSystem.sapSystemNAME = CStr(tstObject)
			Case "SAPGATEWAY"
				If Not IsDBNull(tstObject) Then
					mitaSystem.sapSystemSAPGATEWAY = CStr(tstObject)
				Else
					mitaSystem.sapSystemSAPGATEWAY = ""
				End If
			Case "SAPSERVICE"
				If Not IsDBNull(tstObject) Then
					mitaSystem.sapSystemSAPSERVICE = CStr(tstObject)
				Else
					mitaSystem.sapSystemSAPSERVICE = ""
				End If
			Case "SAPID"
				mitaSystem.sapSystemSAPID = CStr(tstObject)
			Case "SAPIDCLT"
				mitaSystem.sapSystemSAPIDCLT = CStr(tstObject)
			Case "SAPUSER"
				mitaSystem.sapSystemSAPUSER = CStr(tstObject)
			Case "SAPOWNER"
				If Asc(Left$(tstObject.ToString, 1)) = 255 Then
					mitaSystem.sapSystemSAPOWNER = DoXor(CStr(tstObject))
				Else
					mitaSystem.sapSystemSAPOWNER = CStr(tstObject)
				End If
			Case "SAPSYSTEM"
				mitaSystem.sapSystemSAPSYSTEM = CStr(tstObject)
			Case "SAPSERVER"
				mitaSystem.sapSystemSAPSERVER = CStr(tstObject)
			Case "SAPCLIENT"
				mitaSystem.sapSystemSAPCLIENT = CStr(tstObject)
			Case "DATABAS", "DATABASE"
				mitaSystem.sapSystemDATABASE = CStr(tstObject)
			Case "DATABASTYPE", "DATABASETYPE"
				mitaSystem.sapSystemDATABASETYPE = CStr(tstObject)
			Case "DATABASUSER", "DATABASEUSER"
				mitaSystem.sapSystemDATABASEUSER = CStr(tstObject)
			Case "DATABASPWD", "DATABASEPWD"
				If Asc(Left$(tstObject.ToString, 1)) = 255 Then
					mitaSystem.sapSystemDATABASEPWD = DoXor(CStr(tstObject))
				Else
					mitaSystem.sapSystemDATABASEPWD = CStr(tstObject)
				End If
			Case "SAPVERSIONID"
				mitaSystem.sapSystemVERSIONID = CInt(tstObject)
			Case "SAPVERSIONNAME"
				mitaSystem.sapSystemVERSIONNAME = CStr(tstObject)
			Case "SAPSUBVERSION"
				If Not IsDBNull(tstObject) Then
					mitaSystem.sapSystemSUBVERSION = CStr(tstObject)
				Else
					mitaSystem.sapSystemSUBVERSION = ""
				End If
			Case "ENVIRONMENT"
				mitaSystem.runType = CStr(tstObject)
			Case Else
				nam = nam
		End Select
	End Sub
	Public Function readDBSapSystemFromName(ByVal sapName As String, Optional ByVal connString As String = Nothing) As Boolean
		Dim query As String
		Dim result As Boolean
		Dim sb As New System.Text.StringBuilder(1000)
		sb.Append("SELECT t2.sapversionname, t2.sapsubversion, t1.* FROM pscsapsystems t1, pscsapversions t2")
		sb.Append(" WHERE t1.sapname = '" & sapName & "'")
		sb.Append(" AND t2.activ = 'Y'")
		sb.Append(" AND t1.activ = 'Y'")
		sb.Append(" AND t1.sapversionid = t2.sapversionid")
		query = sb.ToString()
		Return readDBSap(query, connString)
	End Function

End Class