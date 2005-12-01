Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports pscMitaDef.CMitaDef
Module modMitaSub

	Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
	Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
	Public Enum workModi
		inputDirectory
		inputSap
		inputDB
	End Enum

	Public Structure startupStructure
		Dim myID As String
		Dim orderDll As PSCMitaOrder.CMitaOrder
		Dim errorDll As PscMitaError.CMitaError
		Dim workModus As workModi
		Dim cmd As String
		Dim user As String
		Dim pwd As String
		Dim dataBase As String
		Dim id As Integer
		Dim sapSystemName As String
		Dim sapSystemID As Integer
		Dim sapVersionName As String
		Dim runType As String
		Dim alive As Integer
		Dim garbageRemove As Boolean
		Dim sameVersion As Boolean
		Dim sqlClass As mitaSqlClass
		Dim orderDirectory As String
		Dim miniSAP As Boolean
		Dim maxList As Integer
		Dim hostName As String
		Dim typChar As String
		Dim formCaption As String
		Dim trace As Char
		Dim pool As String
		Dim iconIndex As Integer
		Dim caption As String
		Dim processID As Integer
		Dim profile As Boolean
	End Structure

	Public startupInfo As startupStructure
	Public odbc_connection As New OdbcConnection
	Public oracleTrans As OdbcTransaction
	Public doProfile As Boolean = False
	Public myProfile As New pscProfilEx.ProfilEx

	Public listAll As New ListBuffer
	Public listWarnings As New ListBuffer
	Public listErrors As New ListBuffer
	Public listLogs As New ListBuffer
	Public processStopped As Boolean
	Public SapOrder As PSCMitaOrder.CMitaOrder
	Public SapError As PscMitaError.CMitaError
	Public savColor(4) As pscLedEx.pscLed.mainColor
	Public savColorIndex As Integer = -1

	Public SAPClient As Short
	Public abortOrder As Boolean
	Public referenceFlag As Boolean

	Public myID As Short
	Public mitaApplication As String

	Public isReadOnly As Boolean
	Public lockingID As Integer
	Public miniSAP As Boolean

	Public EndIt As Boolean
	Public Formloaded As Boolean
	Public actLogTyp As String
	Public listStopped As Boolean
	Public connStr As String
	Public mitaShared As pscMitaShared.CMitaShared
	Public mitaData As pscMitaData.CMitaData
	Public mitaConnect As pscMitaConnect.CMitaConnect
	Public mitaMessage As pscMitaMsg.CMitaMsg

	Public Function startup() As Boolean
		Dim x() As String
		Dim y() As String
		Dim i As Integer
		Dim result As Boolean
		Dim a As String
		Dim saveIt As Boolean
		mitaApplication = "PSC\" & startupInfo.myID
		startupInfo.user = GetSetting("PSC\pscMitaLogin", "lastUser", "login")
		startupInfo.dataBase = GetSetting("PSC\pscMitaLogin", "lastUser", "database")
		startupInfo.pwd = mitaShared.DoXor(GetSetting("PSC\pscMitaLogin", "lastUser", "pwd"))
		x = Split(startupInfo.cmd, "-")
		For i = 0 To UBound(x)
			If x(i) <> "" Then
				y = Split(x(i), " ")
				Select Case UCase(y(0))
					Case "UID", "U"
						startupInfo.user = y(1)
					Case "DSN", "D"
						startupInfo.dataBase = y(1)
					Case "PWD", "P"
						startupInfo.pwd = mitaShared.DoXor(y(1))
				End Select
			End If
		Next i
		If startupInfo.user = "" Or startupInfo.pwd = "" Or startupInfo.dataBase = "" Then
			Dim frmLoginInst As New pscMitaLogin.CLogin
			frmLoginInst.txtUserName.Text = startupInfo.user
			frmLoginInst.txtPassword.Text = startupInfo.pwd
			frmLoginInst.txtBase.Text = startupInfo.dataBase
			If frmLoginInst.ShowDialog() = DialogResult.Cancel Then End
			a = connectionTest(frmLoginInst.connectString)
			If a <> "" Then
				MsgBox(a, MsgBoxStyle.Critical Or MsgBoxStyle.OKOnly)
				frmLoginInst.Dispose()
				End
			End If
			startupInfo.user = frmLoginInst.txtUserName.Text
			startupInfo.dataBase = frmLoginInst.txtBase.Text
			startupInfo.pwd = frmLoginInst.txtPassword.Text
			frmLoginInst.Dispose()
		End If
		connStr = "DSN=" & startupInfo.dataBase & ";UID=" & startupInfo.user & ";PWD=" & startupInfo.pwd & ";"
		startupInfo.trace = CChar(GetSetting(mitaApplication, "Settings", "trace", "0"))
		startupInfo.sapSystemName = GetSetting(mitaApplication, "Settings", "sapSystemName", "PSC")
		startupInfo.sqlClass = CInt(GetSetting(mitaApplication, "Settings", "sqlClass", mitaSqlClass.classAd + mitaSqlClass.classOrder + mitaSqlClass.classInternal))
		startupInfo.alive = CInt(GetSetting(mitaApplication, "Settings", "alive", "10"))
		'startupInfo.runType = GetSetting(mitaApplication, "Settings", "runType", startupInfo.runType)
		startupInfo.miniSAP = CBool(GetSetting(mitaApplication, "Settings", "MiniSAP", "False"))
		'If InStr(startupInfo.runType, "_") = 0 Then startupInfo.runType = startupInfo.typChar & startupInfo.runType
		startupInfo.garbageRemove = CBool(GetSetting(mitaApplication, "Settings", "sameVersion", "True"))
		startupInfo.sameVersion = CBool(GetSetting(mitaApplication, "Settings", "garbageRemove", "True"))
		startupInfo.maxList = CInt(GetSetting(mitaApplication, "Settings", "maxList", "100"))
		startupInfo.pool = GetSetting(mitaApplication, "Settings", "pool", "sap2atex")
		startupInfo.profile = GetSetting(mitaApplication, "Settings", "profile", "False")
		Dim hostName As String = System.Net.Dns.GetHostName
		startupInfo.hostName = hostName
		For i = 0 To UBound(x)
			If x(i) <> "" Then
				x(i) = Replace(x(i), "=", " ")
				While InStr(x(i), "  ") > 0
					x(i) = Replace(x(i), "  ", " ")
				End While
				y = Split(x(i), " ")
				Select Case UCase(y(0))
					Case "ALIVE"
						startupInfo.alive = CInt(y(1))
					Case "SYS", "S"
						startupInfo.sapSystemName = y(1)
					Case "DIR"
						startupInfo.orderDirectory = y(1)
						If startupInfo.orderDirectory = "" Then startupInfo.orderDirectory = Nothing
						If Not IsNothing(startupInfo.orderDirectory) Then startupInfo.workModus = workModi.inputDirectory
					Case "LIST"
						startupInfo.maxList = CInt(y(1))
					Case "LOG", "L"
						startupInfo.sqlClass = 0
						If InStr(1, y(1), "E", CompareMethod.Text) > 0 Then startupInfo.sqlClass = startupInfo.sqlClass Or mitaSqlClass.classError
						If InStr(1, y(1), "A", CompareMethod.Text) > 0 Then startupInfo.sqlClass = startupInfo.sqlClass Or mitaSqlClass.classAd
						If InStr(1, y(1), "O", CompareMethod.Text) > 0 Then startupInfo.sqlClass = startupInfo.sqlClass Or mitaSqlClass.classOrder
						If InStr(1, y(1), "I", CompareMethod.Text) > 0 Then startupInfo.sqlClass = startupInfo.sqlClass Or mitaSqlClass.classInternal
					Case "SAMEVERSION", "SAME"
						startupInfo.sameVersion = mitaShared.getBool(y)
					Case "GARBAGE", "G"
						startupInfo.garbageRemove = mitaShared.getBool(y)
					Case "PROD", "TEST", "DEVELOP"
						startupInfo.runType = startupInfo.typChar & UCase(y(0))
					Case "MINISAP"
						startupInfo.miniSAP = mitaShared.getBool(y)
					Case "PROFILE"
						doProfile = mitaShared.getBool(y)
					Case "SAVE"
						saveIt = True
					Case "TRACE"
						startupInfo.trace = CChar(UCase(y(1)))
					Case "POOL"
						startupInfo.pool = UCase(y(1))
				End Select
			End If
		Next i
		mitaConnect.dbConnectString = connStr
		mitaData.isMiniSAP = startupInfo.miniSAP
		mitaData.mitaApplication = stringForDb(Trim(startupInfo.myID))
		mitaData.createUser = Environment.UserName.ToString
		mitaData.sapPool = startupInfo.pool
		mitaData.sapSystemNAME = startupInfo.sapSystemName
		SapOrder.connectSet = mitaConnect
		SapOrder.sharedSet = mitaShared
		SapOrder.dataSet = mitaData
		SapError.connectSet = mitaConnect
		SapError.sharedSet = mitaShared
		SapError.dataSet = mitaData
		SapOrder.namesSetSapError(SapError)
		result = startupInfo.orderDll.optionsSetSapTrace(startupInfo.trace)
		If result Then result = result And startupInfo.orderDll.optionsSetGarbageRemove(startupInfo.garbageRemove)
		If result Then result = result And startupInfo.orderDll.optionsSetSqlClass(startupInfo.sqlClass)
		If result Then
			If saveIt Then
				SaveSetting(mitaApplication, "Settings", "sapSystemName", startupInfo.sapSystemName)
				SaveSetting(mitaApplication, "Settings", "sqlClass", CStr(startupInfo.sqlClass))
				SaveSetting(mitaApplication, "Settings", "garbageRemove", CStr(startupInfo.garbageRemove))
				SaveSetting(mitaApplication, "Settings", "sameVersion", CStr(startupInfo.sameVersion))
				SaveSetting(mitaApplication, "Settings", "miniSap", CStr(startupInfo.miniSAP))
				SaveSetting(mitaApplication, "Settings", "alive", CStr(startupInfo.alive))
				SaveSetting(mitaApplication, "Settings", "maxList", CStr(startupInfo.maxList))
				SaveSetting(mitaApplication, "Settings", "pool", startupInfo.pool)
				SaveSetting(mitaApplication, "Settings", "trace", startupInfo.trace)
				SaveSetting(mitaApplication, "Settings", "profile", startupInfo.profile)
			End If
			SaveSetting("PSC\pscMitaLogin", "lastUser", "login", startupInfo.user)
			SaveSetting("PSC\pscMitaLogin", "lastUser", "database", startupInfo.dataBase)
			SaveSetting("PSC\pscMitaLogin", "lastUser", "pwd", mitaShared.DoXor(startupInfo.pwd))
			Dim cmd As String = "/SYS=" & startupInfo.sapSystemName
			cmd = cmd & " /LOG="
			If startupInfo.sqlClass And mitaSqlClass.classAd = mitaSqlClass.classAd Then cmd = cmd & "A"
			If startupInfo.sqlClass And mitaSqlClass.classError = mitaSqlClass.classError Then cmd = cmd & "E"
			If startupInfo.sqlClass And mitaSqlClass.classInternal = mitaSqlClass.classInternal Then cmd = cmd & "I"
			If startupInfo.sqlClass And mitaSqlClass.classOrder = mitaSqlClass.classOrder Then cmd = cmd & "O"
			cmd = cmd & " /" & Mid$(startupInfo.runType, 3)
			cmd = cmd & " /SAMEVERSION=" & IIf(startupInfo.sameVersion, "On", "Off")
			cmd = cmd & " /GARBAGE=" & IIf(startupInfo.garbageRemove, "On", "Off")
			cmd = cmd & " /MINISAP=" & IIf(startupInfo.miniSAP, "On", "Off")
			cmd = cmd & " /PROFILE=" & IIf(doProfile, "On", "Off")
			cmd = cmd & " /POOL=" & startupInfo.pool
			cmd = cmd & " /TRACE=" & startupInfo.trace
			cmd = cmd & " /LIST=" & startupInfo.maxList
			If startupInfo.workModus = workModi.inputDirectory Then
				cmd = cmd & " /DIR=" & startupInfo.orderDirectory
			End If
			startupInfo.cmd = cmd
			mitaData.commandLine = startupInfo.cmd
			If doProfile Then
				myProfile.profileOpen(Application.StartupPath & "\ProfilOutput\" & startupInfo.myID)
			End If
			With odbc_connection
				.ConnectionString = connStr
				.ConnectionTimeout = 5
			End With
			startupInfo.sapSystemID = mitaShared.getSapIDFromName(odbc_connection, startupInfo.sapSystemName)
			If startupInfo.runType = "" Then
				mitaData.runType = startupInfo.typChar & mitaData.runType
			Else
				mitaData.runType = startupInfo.runType
			End If
		End If
		Return result
	End Function
	Public Sub Pause(ByRef TenthOfSec As Integer)
		Dim i As Integer
		On Error Resume Next
		For i = 1 To TenthOfSec
			Sleep(99)
			System.Windows.Forms.Application.DoEvents()
		Next i
	End Sub
	Public Function ShowCaption(ByRef applicationForm As System.Windows.Forms.Form) As Short
		Dim C As String
		Dim t As String
		Dim w As Integer
		Dim z As Short
		Dim p As Short
		C = applicationForm.Text
		t = C & " #1"
		p = Len(t)
		t = t & "  -  Version " & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart & "." & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMinorPart
		Dim tmp As New System.Windows.Forms.Form
		tmp.Text = t
		mitaShared.buildVersionInfo(tmp)
		t = tmp.Text
		tmp.Dispose()
		z = 1
		Do
			Mid(t, p, 1) = Trim(Str(z))
			w = FindWindow(vbNullString, t)
			If w = 0 Then Exit Do
			z = z + 1
		Loop
		applicationForm.Text = t
		ShowCaption = z
	End Function
	Public Sub colorStop()
		If Not Formloaded Then Exit Sub
		savColorIndex = -1
		applicationForm.processLed.ledColor = pscLedEx.pscLed.mainColor.colorRed
		System.Windows.Forms.Application.DoEvents()
	End Sub

	Public Sub colorSap()
		If Not Formloaded Then Exit Sub
		savColorIndex = savColorIndex + 1
		savColor(savColorIndex) = applicationForm.processLed.ledColor
		applicationForm.processLed.ledColor = pscLedEx.pscLed.mainColor.colorYellow
		System.Windows.Forms.Application.DoEvents()
	End Sub
	Public Sub colorDB()
		If Not Formloaded Then Exit Sub
		savColorIndex = savColorIndex + 1
		savColor(savColorIndex) = applicationForm.processLed.ledColor
		applicationForm.processLed.ledColor = pscLedEx.pscLed.mainColor.colorCyan
		System.Windows.Forms.Application.DoEvents()
	End Sub
	Public Sub colorProcess()
		If Not Formloaded Then Exit Sub
		savColorIndex = savColorIndex + 1
		savColor(savColorIndex) = applicationForm.processLed.ledColor
		applicationForm.processLed.ledColor = pscLedEx.pscLed.mainColor.colorOrange
		System.Windows.Forms.Application.DoEvents()
	End Sub
	Public Sub colorIdle()
		If Not Formloaded Then Exit Sub
		savColorIndex = -1
		applicationForm.processLed.ledColor = pscLedEx.pscLed.mainColor.colorGrey
		System.Windows.Forms.Application.DoEvents()
	End Sub
	Public Sub colorRestore()
		If Not Formloaded Then Exit Sub
		applicationForm.processLed.ledColor = savColor(savColorIndex)
		savColorIndex = savColorIndex - 1
		System.Windows.Forms.Application.DoEvents()
	End Sub
	Public Sub listBoxToFile(ByRef lst As System.Windows.Forms.ListBox, ByRef subDir As String, ByRef fileName As String, ByVal descending As Boolean)
		Dim timeString As String
		Dim f1 As Integer
		Dim filNam As String
		Dim applicationName As String
		Dim path As String
		Dim i As Integer
		timeString = Format(Now, "yyyyMMddHHmmss")
		path = Application.ExecutablePath
		i = InStrRev(path, "\")
		applicationName = Mid$(path, i + 1)
		path = VB.Left$(path, i) & subDir
		i = InStr(applicationName, ".")
		applicationName = VB.Left$(applicationName, i - 1)
		filNam = fileName & "_" & timeString & ".log"
		If Dir$(path, vbDirectory) = "" Then
			MkDir(path)
		End If
		path = path & "\"
		If Dir$(path & filNam) <> "" Then Kill(path & filNam)
		f1 = VB.FreeFile()
		VB.FileOpen(f1, path & filNam, VB.OpenMode.Append)
		VB.PrintLine(f1, "Log File Copy (from " & lst.Name & ") of: SAP System '" & startupInfo.sapSystemName & "' on Host '" & startupInfo.hostName & "', ID " & applicationForm.myInstance)
		VB.PrintLine(f1, "Created by pscMitaOLP at: " & CStr(Now))
		VB.PrintLine(f1, "")
		If descending Then
			For i = lst.Items.Count - 1 To 0 Step -1
				VB.PrintLine(f1, lst.Items.Item(i).ToString)
			Next
		Else
			For i = 0 To lst.Items.Count - 1
				VB.PrintLine(f1, lst.Items.Item(i).ToString)
			Next
		End If
		VB.FileClose(f1)
	End Sub
	Public Sub fillListBox(ByRef lst As System.Windows.Forms.ListBox, ByRef src As ListBuffer)
		Dim newList() As String = src.readItems
		lst.Items.Clear()
		If IsNothing(newList) Then Exit Sub
		Dim i As Integer
		For i = UBound(newList) To 0 Step -1
			lst.Items.Add(newList(i))
		Next
	End Sub
	Public Function connectionTest(ByRef connectString As String) As String
		Dim conn As New OdbcConnection
		Try
			With conn
				.ConnectionString = connectString
				.Open()
			End With
			connectionTest = ""
			conn.Close()
		Catch
			connectionTest = Err.Description
		Finally
			conn.Dispose()
			conn = Nothing
		End Try
	End Function
	Public Sub deleteList(ByVal lst As Windows.Forms.ListBox)
		While lst.Items.Count > startupInfo.maxList
			lst.Items.Remove(lst.Items(lst.Items.Count - 1))
		End While
	End Sub
	Public Function dirProcess() As Boolean
		Dim a As String
		Dim orders() As String
		Dim orderCnt As Integer
		Dim o As Integer
		Dim dummy As Integer
		Dim i As Integer
		Dim path As String
		Dim ext As String
		Dim startTime As Date
		startTime = System.DateTime.FromOADate(VB.Timer())
		SapOrder.optionsSetGarbageRemove(startupInfo.garbageRemove)
		SapOrder.optionsAllowSameVersion(startupInfo.sameVersion)
		i = InStrRev(startupInfo.orderDirectory, "\")
		path = VB.Left$(startupInfo.orderDirectory, i)
		ext = VB.Mid$(startupInfo.orderDirectory, i + 1)
		a = Dir(path & ext)
		dirProcess = False
		If a = "" Then Exit Function
		orderCnt = -1
		If VB.Left(a, 1) <> "." Then
			orderCnt = orderCnt + 1
			ReDim Preserve orders(orderCnt)
			orders(orderCnt) = a
			Do
				a = Dir$()
				If a = "" Then Exit Do
				If VB.Left(a, 1) <> "." Then
					orderCnt = orderCnt + 1
					ReDim Preserve orders(orderCnt)
					orders(orderCnt) = a
				End If
				If EndIt Or processStopped Then Exit Do
			Loop
		End If
		For o = 0 To orderCnt
			abortOrder = False
			If SapOrder.orderReadFile(path & orders(o)) Then
				If EndIt Or processStopped Then Exit For
				If Not abortOrder Then doProcess()
				SapOrder.orderAbort()
			End If
			If EndIt Or processStopped Then Exit For
		Next o
		Dim e As New System.Windows.Forms.ToolBarButtonClickEventArgs(applicationForm.ToolBarButton11)
		applicationForm.ToolBar1.Buttons.Item(10).Pushed = True
		applicationForm.ToolBar1_ButtonClick(Nothing, e)
		dirProcess = True
	End Function

	Public Function dbProcess() As Boolean
		Dim dummy As Long
		Dim hasNext As Boolean
		hasNext = SapOrder.orderReadNextDB()
		If Not hasNext Then
			Return False
		Else
			doProcess()
		End If
		Return True
	End Function
	Public Function sapProcess() As Boolean
		Dim dummy As Long
		Dim hasNext As Boolean
		hasNext = SapOrder.orderReadSap
		If Not hasNext Then
			Sleep(300)
			Return False
		Else
			If SapOrder.orderTransfer() Then doProcess()
		End If
		Return True
	End Function
	Public Sub showApplication()
		startupInfo.cmd = UCase(Replace(VB.Command(), "/", "-"))
		If Not startup() Then
			applicationForm.Close()
			End
		End If
		applicationForm.Label2.Text = startupInfo.cmd
		applicationForm.ToolTip1.SetToolTip(applicationForm.Label2, startupInfo.cmd)
		applicationForm.pscOrder = startupInfo.orderDll
		applicationForm.pscError = startupInfo.errorDll
		listAll.setMaxList(startupInfo.maxList)
		listWarnings.setMaxList(startupInfo.maxList)
		listErrors.setMaxList(startupInfo.maxList)
		listLogs.setMaxList(startupInfo.maxList)
	End Sub
	Public Sub generateDLLs()
		mitaData = New pscMitaData.CMitaData
		mitaData.rfcType = rfcClass.rfcOrder
		mitaConnect = New pscMitaConnect.CMitaConnect
		mitaShared = New pscMitaShared.CMitaShared
		SapOrder = New PSCMitaOrder.CMitaOrder
		SapError = New PscMitaError.CMitaError
		mitaMessage = New pscMitaMsg.CMitaMsg
		mitaData.frmMitaMsgInst = mitaMessage
		SapOrder.optionsSetSqlClass((mitaSqlClass.classAd + mitaSqlClass.classOrder))
		startupInfo.orderDll = SapOrder
		startupInfo.errorDll = SapError
		startupInfo.myID = mitaShared.getExeName()
	End Sub
	Public Function stringForDb(ByRef buf1 As String) As String
		Dim buf2 As String
		buf2 = Replace(buf1, "'", " ")
		buf2 = Replace(buf2, "@", " ")
		buf2 = Replace(buf2, "´", " ")
		buf2 = Replace(buf2, "`", " ")
		stringForDb = buf2
	End Function
End Module