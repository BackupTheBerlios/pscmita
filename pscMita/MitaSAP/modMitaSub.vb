Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports pscMitaDef.CMitaDef
Module modMitaSub

	Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
	Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
	Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer

	Public odbc_connection As New OdbcConnection
	Public oracleTrans As OdbcTransaction
	Public doProfile As Boolean = False
	Public myProfile As New pscProfilEx.CProfilEx

	Public listAll As New ListBuffer
	Public listWarnings As New ListBuffer
	Public listErrors As New ListBuffer
	Public listLogs As New ListBuffer
	Public processStopped As Boolean
	Public isIdle As Boolean
	Public SapOrder As PSCMitaOrder.CMitaOrder
	Public SapError As PscMitaError.CMitaError
	Public savColor(4) As pscLedEx.pscLed.mainColor
	Public savColorIndex As Integer = -1

	Public SAPClient As Short
	Public abortOrder As Boolean
	'Public referenceFlag As Boolean

	Public myID As Short
	Public mitaApplication As String

	Public isReadOnly As Boolean
	Public lockingID As Integer
	Public miniSAP As Boolean

	Public EndIt As Boolean
	Public Formloaded As Boolean
	Public actLogTyp As String
	Public listStopped As Boolean
	Public mitaShared As pscMitaShared.CMitaShared
	Public mitaData As pscMitaData.CMitaData
	Public mitaSystem As New pscMitaSapSystem.CMitaSapSystem
	Public mitaMessage As pscMitaMsg.CMitaMsg
	Public mitaConnect As pscMitaConnect.CMitaConnect
	Public connStr As String

	Public Function startup() As Boolean
		Dim i As Integer
		Dim result As Boolean
		Dim a As String
		Dim saveIt As Boolean
		Dim x() As String
		Dim y() As String
		startupInfo.hideMe = defaultInfo.hideMe
		startupInfo.centerMe = defaultInfo.centerMe
		mitaApplication = "PSC\" & startupInfo.myID
		Dim loginCls As New pscMitaLogin.CLogin
		loginCls.registryKey = "PSC\pscMitaLogin"
		If Not loginCls.doLogin Then End
		connStr = loginCls.connectString
		loginCls.Dispose()
		mitaConnect.dbConnectString = connStr
		startupInfo.sapSystemName = GetSetting(mitaApplication, "Settings", "sapSystemName", "")
		If startupInfo.sapSystemName = "" Then
			startupInfo.sapSystemName = "PSC"
			saveIt = True
		End If
		If usedInfo.trace Then startupInfo.trace = CChar(GetSetting(mitaApplication, "Settings", "trace", "0"))
		If usedInfo.miniSAP Then startupInfo.miniSAP = CBool(GetSetting(mitaApplication, "Settings", "MiniSAP", defaultInfo.miniSAP))
		If usedInfo.profile Then startupInfo.profile = CBool(GetSetting(mitaApplication, "Settings", "profile", defaultInfo.profile))
		If usedInfo.sqlClass Then startupInfo.sqlClass = CInt(GetSetting(mitaApplication, "Settings", "sqlClass", defaultInfo.sqlClass))
		If usedInfo.alive Then startupInfo.alive = CInt(GetSetting(mitaApplication, "Settings", "alive", CStr(defaultInfo.alive)))
		If usedInfo.garbageRemove Then startupInfo.garbageRemove = CBool(GetSetting(mitaApplication, "Settings", "garbageRemove", defaultInfo.garbageRemove))
		If usedInfo.sameVersion Then startupInfo.sameVersion = CBool(GetSetting(mitaApplication, "Settings", "sameVersion", defaultInfo.sameVersion))
		If usedInfo.maxList Then startupInfo.maxList = CInt(GetSetting(mitaApplication, "Settings", "maxList", CStr(defaultInfo.maxList)))
		If usedInfo.pool Then startupInfo.pool = GetSetting(mitaApplication, "Settings", "pool", defaultInfo.pool)
		If usedInfo.runType Then startupInfo.runType = GetSetting(mitaApplication, "Settings", "runType", defaultInfo.runType)
		If InStr(startupInfo.runType, "_") <> 0 Then
			startupInfo.runType = Mid$(startupInfo.runType, 3)
		End If
		If usedInfo.taskBar Then startupInfo.taskBar = CBool(GetSetting(mitaApplication, "Settings", "taskBar", defaultInfo.taskBar))
		If usedInfo.iconTray Then startupInfo.iconTray = CBool(GetSetting(mitaApplication, "Settings", "iconTray", defaultInfo.iconTray))
		If usedInfo.hideMe Then startupInfo.hideMe = CBool(GetSetting(mitaApplication, "Settings", "hideMe", defaultInfo.hideMe))
		If usedInfo.centerMe Then startupInfo.centerMe = CBool(GetSetting(mitaApplication, "Settings", "centerMe", defaultInfo.centerMe))
		If usedInfo.batchError Then startupInfo.batchError = CBool(GetSetting(mitaApplication, "Settings", "batchError", defaultInfo.batchError))
		startupInfo.hostName = System.Net.Dns.GetHostName
		x = Split(startupInfo.cmd, "-")
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
						startupInfo.runType = UCase(y(0))
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
					Case "TASKBAR"
						startupInfo.taskBar = mitaShared.getBool(y)
					Case "ICONTRAY"
						startupInfo.iconTray = mitaShared.getBool(y)
					Case "RESTORE"
						startupInfo.hideMe = False
						startupInfo.centerMe = False
					Case "CENTER"
						startupInfo.hideMe = False
						startupInfo.centerMe = True
					Case "HIDE"
						startupInfo.hideMe = True
						startupInfo.centerMe = False
					Case "ERRBATCH"
						startupInfo.batchError = mitaShared.getBool(y)
				End Select
			End If
		Next i
		mitaData.isMiniSAP = startupInfo.miniSAP
		mitaData.mitaApplication = stringForDb(Trim(startupInfo.myID))
		mitaData.createUser = Environment.UserName.ToString
		mitaSystem.sapPool = startupInfo.pool
		mitaShared.dataSet = mitaData
		mitaShared.readDBSapSystemFromName(startupInfo.sapSystemName)
		SapOrder.dataSet = mitaData
		SapOrder.connectSet = mitaConnect
		SapOrder.sharedSet = mitaShared
		SapError.dataSet = mitaData
		SapError.connectSet = mitaConnect
		SapError.sharedSet = mitaShared
		SapOrder.namesSetSapError(SapError)
		result = startupInfo.orderDll.optionsSetSapTrace(startupInfo.trace)
		If result Then result = result And startupInfo.orderDll.optionsSetGarbageRemove(startupInfo.garbageRemove)
		If result Then result = result And startupInfo.orderDll.optionsSetSqlClass(startupInfo.sqlClass)
		If result Then
			With odbc_connection
				.ConnectionString = connStr
				.ConnectionTimeout = 5
			End With
			startupInfo.sapSystemID = mitaShared.getSapIDFromName(odbc_connection, startupInfo.sapSystemName)
			If startupInfo.runType = "" Then
				mitaSystem.runType = startupInfo.typChar & mitaSystem.runType
			Else
				mitaSystem.runType = startupInfo.runType
			End If
			startupInfo.runType = mitaSystem.runType
			If saveIt Then
				SaveSetting(mitaApplication, "Settings", "sapSystemName", startupInfo.sapSystemName)
				If usedInfo.sqlClass Then SaveSetting(mitaApplication, "Settings", "sqlClass", CStr(startupInfo.sqlClass))
				If usedInfo.garbageRemove Then SaveSetting(mitaApplication, "Settings", "garbageRemove", CStr(startupInfo.garbageRemove))
				If usedInfo.sameVersion Then SaveSetting(mitaApplication, "Settings", "sameVersion", CStr(startupInfo.sameVersion))
				If usedInfo.miniSAP Then SaveSetting(mitaApplication, "Settings", "miniSap", CStr(startupInfo.miniSAP))
				If usedInfo.alive Then SaveSetting(mitaApplication, "Settings", "alive", CStr(startupInfo.alive))
				If usedInfo.maxList Then SaveSetting(mitaApplication, "Settings", "maxList", CStr(startupInfo.maxList))
				If usedInfo.pool Then SaveSetting(mitaApplication, "Settings", "pool", startupInfo.pool)
				If usedInfo.runType Then SaveSetting(mitaApplication, "Settings", "runType", startupInfo.runType)
				If usedInfo.trace Then SaveSetting(mitaApplication, "Settings", "trace", startupInfo.trace)
				If usedInfo.profile Then SaveSetting(mitaApplication, "Settings", "profile", startupInfo.profile)
				If usedInfo.taskBar Then SaveSetting(mitaApplication, "Settings", "taskBar", CStr(startupInfo.taskBar))
				If usedInfo.iconTray Then SaveSetting(mitaApplication, "Settings", "iconTray", CStr(startupInfo.iconTray))
				If usedInfo.hideMe Then SaveSetting(mitaApplication, "Settings", "hideMe", CStr(startupInfo.hideMe))
				If usedInfo.centerMe Then SaveSetting(mitaApplication, "Settings", "centerMe", CStr(startupInfo.centerMe))
				If usedInfo.batchError Then SaveSetting(mitaApplication, "Settings", "batchError", CStr(startupInfo.batchError))
			End If
			Dim cmd As String = "/SYS=" & startupInfo.sapSystemName
			If usedInfo.sqlClass And (startupInfo.sqlClass <> defaultInfo.sqlClass) Then
				cmd = cmd & " /LOG="
				If (startupInfo.sqlClass And mitaSqlClass.classAd) = mitaSqlClass.classAd Then cmd = cmd & "A"
				If (startupInfo.sqlClass And mitaSqlClass.classError) = mitaSqlClass.classError Then cmd = cmd & "E"
				If (startupInfo.sqlClass And mitaSqlClass.classInternal) = mitaSqlClass.classInternal Then cmd = cmd & "I"
				If (startupInfo.sqlClass And mitaSqlClass.classOrder) = mitaSqlClass.classOrder Then cmd = cmd & "O"
			End If
			If usedInfo.runType And (Not startupInfo.runType.EndsWith(defaultInfo.runType)) Then cmd = cmd & " /" & startupInfo.runType
			If usedInfo.taskBar And (startupInfo.taskBar <> defaultInfo.taskBar) Then cmd = cmd & " /TASKBAR=" & IIf(startupInfo.taskBar, "On", "Off")
			If usedInfo.iconTray And (startupInfo.iconTray <> defaultInfo.iconTray) Then cmd = cmd & " /ICONTRAY=" & IIf(startupInfo.iconTray, "On", "Off")
			If usedInfo.hideMe And (startupInfo.hideMe <> defaultInfo.hideMe) Then
				cmd = cmd & " /HIDE"
			ElseIf usedInfo.centerMe And (startupInfo.centerMe <> defaultInfo.centerMe) Then
				cmd = cmd & " /CENTER"
			End If
			If usedInfo.alive And (startupInfo.alive <> defaultInfo.alive) Then cmd = cmd & " /ALIVE=" & CStr(startupInfo.alive)
			If usedInfo.sameVersion And (startupInfo.sameVersion <> defaultInfo.sameVersion) Then cmd = cmd & " /SAMEVERSION=" & IIf(startupInfo.sameVersion, "On", "Off")
			If usedInfo.garbageRemove And (startupInfo.garbageRemove <> defaultInfo.garbageRemove) Then cmd = cmd & " /GARBAGE=" & IIf(startupInfo.garbageRemove, "On", "Off")
			If usedInfo.miniSAP And (startupInfo.miniSAP <> defaultInfo.miniSAP) Then cmd = cmd & " /MINISAP=" & IIf(startupInfo.miniSAP, "On", "Off")
			If usedInfo.batchError And (startupInfo.batchError <> defaultInfo.batchError) Then cmd = cmd & " /ERRBATCH=" & IIf(startupInfo.batchError, "On", "Off")
			If usedInfo.profile And (startupInfo.profile <> defaultInfo.profile) Then cmd = cmd & " /PROFILE=" & IIf(doProfile, "On", "Off")
			If usedInfo.pool And (startupInfo.pool <> defaultInfo.pool) Then cmd = cmd & " /POOL=" & startupInfo.pool
			If usedInfo.trace And (startupInfo.trace <> defaultInfo.trace) Then cmd = cmd & " /TRACE=" & startupInfo.trace
			If usedInfo.maxList And (startupInfo.maxList <> defaultInfo.maxList) Then cmd = cmd & " /LIST=" & startupInfo.maxList
			If usedInfo.workModus And (startupInfo.workModus = workModi.inputDirectory) Then
				cmd = cmd & " /DIR=" & startupInfo.orderDirectory
			End If
			startupInfo.cmd = cmd
			mitaData.commandLine = startupInfo.cmd
			If doProfile Then
				myProfile.profileOpen(Application.StartupPath & "\ProfilOutput\" & startupInfo.myID)
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
		If savColorIndex = -1 Then Exit Sub
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
		VB.PrintLine(f1, "Log File Copy (from " & lst.Name & ") of: SAP System '" & startupInfo.sapSystemName & "' on Host '" & startupInfo.hostName & "', ID " & myInstance)
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
		mitaSystem.rfcType = rfcClass.rfcOrder
		mitaConnect = New pscMitaConnect.CMitaConnect
		mitaShared = New pscMitaShared.CMitaShared
		SapOrder = New PSCMitaOrder.CMitaOrder
		SapError = New PscMitaError.CMitaError
		mitaMessage = New pscMitaMsg.CMitaMsg
		SapOrder.messageSet = mitaMessage
		SapOrder.systemSet = mitaSystem
		SapError.systemSet = mitaSystem
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
	Public Sub setDefault(ByVal field As String, ByVal content As Object)
		Select Case field
			Case "runType"
				defaultInfo.runType = content.ToString
				usedInfo.runType = True
			Case "alive"
				defaultInfo.alive = CInt(content)
				usedInfo.alive = True
			Case "batchError"
				defaultInfo.batchError = CBool(content)
				usedInfo.batchError = True
			Case "garbageRemove"
				defaultInfo.garbageRemove = CBool(content)
				usedInfo.garbageRemove = True
			Case "iconTray"
				defaultInfo.iconTray = CBool(content)
				usedInfo.iconTray = True
			Case "maxList"
				defaultInfo.maxList = CInt(content)
				usedInfo.maxList = True
			Case "profile"
				defaultInfo.profile = CBool(content)
				usedInfo.profile = True
			Case "sameVersion"
				defaultInfo.sameVersion = CBool(content)
				usedInfo.sameVersion = True
			Case "sqlClass"
				defaultInfo.sqlClass = CInt(content)
				usedInfo.sqlClass = True
			Case "taskBar"
				defaultInfo.taskBar = CBool(content)
				usedInfo.taskBar = True
			Case "trace"
				defaultInfo.trace = content.ToString
				usedInfo.trace = True
			Case "pool"
				defaultInfo.pool = content.ToString
				usedInfo.pool = True
			Case "hideMe"
				defaultInfo.hideMe = CBool(content)
				usedInfo.hideMe = True
			Case "miniSAP"
				defaultInfo.miniSAP = CBool(content)
				usedInfo.miniSAP = True
			Case "centerMe"
				defaultInfo.centerMe = CBool(content)
				usedInfo.centerMe = True
			Case "workModus"
				defaultInfo.workModus = CInt(content)
				usedInfo.workModus = True
			Case Else
				MsgBox(field)
		End Select
	End Sub
	Public Function getSapField(ByRef fieldName As String) As String
		Dim result As Boolean
		Dim tmp As String
		result = SapOrder.getFieldValue(fieldName, tmp)
		If result Then
			getSapField = tmp
		Else
			getSapField = ""
		End If
	End Function
End Module