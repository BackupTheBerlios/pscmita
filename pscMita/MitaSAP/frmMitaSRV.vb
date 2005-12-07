Option Strict Off
Option Explicit On 
Imports VB = Microsoft.VisualBasic
Imports pscMitaDef.CMitaDef
Imports NPSCLed.PscLed
Friend Class frmMitaSRV
	Inherits System.Windows.Forms.Form

#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents Timer2 As System.Windows.Forms.Timer
	Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
	Friend WithEvents Panel2 As System.Windows.Forms.Panel
	Friend WithEvents ToolBarButton1 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton2 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton3 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton4 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton5 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton6 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton7 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton8 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton9 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBar1 As System.Windows.Forms.ToolBar
	Friend WithEvents ToolBarButton10 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton11 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton12 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton13 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton14 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton15 As System.Windows.Forms.ToolBarButton
	Friend WithEvents ToolBarButton16 As System.Windows.Forms.ToolBarButton
	Friend WithEvents SAPMld As System.Windows.Forms.ListBox
	Friend WithEvents processLed As NPSCLed.PscLed
	Friend WithEvents aliveLed As NPSCLed.PscLed
	Friend WithEvents NotifyIcon As System.Windows.Forms.NotifyIcon
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMitaSRV))
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
		Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
		Me.SAPMld = New System.Windows.Forms.ListBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
		Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
		Me.Panel2 = New System.Windows.Forms.Panel
		Me.ToolBar1 = New System.Windows.Forms.ToolBar
		Me.ToolBarButton1 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton5 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton2 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton4 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton3 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton6 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton7 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton8 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton9 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton10 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton11 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton12 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton15 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton14 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton13 = New System.Windows.Forms.ToolBarButton
		Me.ToolBarButton16 = New System.Windows.Forms.ToolBarButton
		Me.processLed = New NPSCLed.PscLed
		Me.aliveLed = New NPSCLed.PscLed
		Me.NotifyIcon = New System.Windows.Forms.NotifyIcon(Me.components)
		Me.Panel2.SuspendLayout()
		Me.SuspendLayout()
		'
		'ToolTip1
		'
		Me.ToolTip1.AutoPopDelay = 10000
		Me.ToolTip1.InitialDelay = 500
		Me.ToolTip1.ReshowDelay = 100
		Me.ToolTip1.ShowAlways = True
		'
		'Timer1
		'
		Me.Timer1.Interval = 1000
		'
		'SAPMld
		'
		Me.SAPMld.BackColor = System.Drawing.SystemColors.Window
		Me.SAPMld.Cursor = System.Windows.Forms.Cursors.Default
		Me.SAPMld.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.SAPMld.ForeColor = System.Drawing.SystemColors.WindowText
		Me.SAPMld.ItemHeight = 14
		Me.SAPMld.Location = New System.Drawing.Point(8, 54)
		Me.SAPMld.Name = "SAPMld"
		Me.SAPMld.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.SAPMld.Size = New System.Drawing.Size(496, 130)
		Me.SAPMld.TabIndex = 1
		'
		'Label2
		'
		Me.Label2.BackColor = System.Drawing.SystemColors.Control
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label2.Location = New System.Drawing.Point(8, 196)
		Me.Label2.Name = "Label2"
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.Size = New System.Drawing.Size(434, 17)
		Me.Label2.TabIndex = 4
		'
		'Timer2
		'
		'
		'ImageList1
		'
		Me.ImageList1.ImageSize = New System.Drawing.Size(33, 33)
		Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
		'
		'Panel2
		'
		Me.Panel2.Controls.Add(Me.ToolBar1)
		Me.Panel2.Location = New System.Drawing.Point(8, 2)
		Me.Panel2.Name = "Panel2"
		Me.Panel2.Size = New System.Drawing.Size(416, 48)
		Me.Panel2.TabIndex = 15
		'
		'ToolBar1
		'
		Me.ToolBar1.Buttons.AddRange(New System.Windows.Forms.ToolBarButton() {Me.ToolBarButton1, Me.ToolBarButton5, Me.ToolBarButton2, Me.ToolBarButton4, Me.ToolBarButton3, Me.ToolBarButton6, Me.ToolBarButton7, Me.ToolBarButton8, Me.ToolBarButton9, Me.ToolBarButton10, Me.ToolBarButton11, Me.ToolBarButton12, Me.ToolBarButton15, Me.ToolBarButton14, Me.ToolBarButton13, Me.ToolBarButton16})
		Me.ToolBar1.ButtonSize = New System.Drawing.Size(33, 33)
		Me.ToolBar1.Divider = False
		Me.ToolBar1.DropDownArrows = True
		Me.ToolBar1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.ToolBar1.ImageList = Me.ImageList1
		Me.ToolBar1.Location = New System.Drawing.Point(0, 0)
		Me.ToolBar1.Name = "ToolBar1"
		Me.ToolBar1.ShowToolTips = True
		Me.ToolBar1.Size = New System.Drawing.Size(416, 43)
		Me.ToolBar1.TabIndex = 0
		'
		'ToolBarButton1
		'
		Me.ToolBarButton1.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
		'
		'ToolBarButton5
		'
		Me.ToolBarButton5.ImageIndex = 0
		Me.ToolBarButton5.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
		Me.ToolBarButton5.ToolTipText = "Show All Logs"
		'
		'ToolBarButton2
		'
		Me.ToolBarButton2.ImageIndex = 2
		Me.ToolBarButton2.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
		Me.ToolBarButton2.ToolTipText = "Show Normal Logs"
		'
		'ToolBarButton4
		'
		Me.ToolBarButton4.ImageIndex = 1
		Me.ToolBarButton4.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
		Me.ToolBarButton4.ToolTipText = "Show Warning Logs"
		'
		'ToolBarButton3
		'
		Me.ToolBarButton3.ImageIndex = 3
		Me.ToolBarButton3.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
		Me.ToolBarButton3.ToolTipText = "Show Error Logs"
		'
		'ToolBarButton6
		'
		Me.ToolBarButton6.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
		'
		'ToolBarButton7
		'
		Me.ToolBarButton7.ImageIndex = 8
		Me.ToolBarButton7.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
		Me.ToolBarButton7.ToolTipText = "Stop List Refresh"
		'
		'ToolBarButton8
		'
		Me.ToolBarButton8.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
		'
		'ToolBarButton9
		'
		Me.ToolBarButton9.ImageIndex = 6
		Me.ToolBarButton9.ToolTipText = "Save List to File"
		'
		'ToolBarButton10
		'
		Me.ToolBarButton10.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
		'
		'ToolBarButton11
		'
		Me.ToolBarButton11.ImageIndex = 5
		Me.ToolBarButton11.Style = System.Windows.Forms.ToolBarButtonStyle.ToggleButton
		Me.ToolBarButton11.ToolTipText = "Stop Processing"
		'
		'ToolBarButton12
		'
		Me.ToolBarButton12.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
		'
		'ToolBarButton15
		'
		Me.ToolBarButton15.ImageIndex = 10
		Me.ToolBarButton15.ToolTipText = "About ..."
		'
		'ToolBarButton14
		'
		Me.ToolBarButton14.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
		'
		'ToolBarButton13
		'
		Me.ToolBarButton13.ImageIndex = 9
		'
		'ToolBarButton16
		'
		Me.ToolBarButton16.Style = System.Windows.Forms.ToolBarButtonStyle.Separator
		'
		'processLed
		'
		Me.processLed.CausesValidation = False
		Me.processLed.ledBackTransparent = False
		Me.processLed.ledBeep = ledAudio.audioOff
		Me.processLed.ledBlinkSpeed = blinkSpeed.blinkVeryFast
		Me.processLed.ledBorder = ledBorderStyle.None
		Me.processLed.ledColor = mainColor.colorGreen
		Me.processLed.ledColorBlink = blinkColor.colorDark
		Me.processLed.ledDesignBehaviour = ledDesignMode.blinkOn_audioOff
		Me.processLed.ledFlashNext = Nothing
		Me.processLed.ledNextTrigger = ledTrigger.nextDirect
		Me.processLed.ledSize = ledModel.sizeLarge
		Me.processLed.ledSoundFile = ""
		Me.processLed.ledStatus = ledModus.ledOn
		Me.processLed.Location = New System.Drawing.Point(448, 186)
		Me.processLed.Name = "processLed"
		Me.processLed.Size = New System.Drawing.Size(32, 32)
		Me.processLed.socketColor = System.Drawing.KnownColor.Silver
		Me.processLed.socketFeature = ledSocketFeature.socketRaised
		Me.processLed.socketWidth = ledSocketWidth.widthSmall
		Me.processLed.TabIndex = 16
		Me.processLed.toolTip = Nothing
		'
		'aliveLed
		'
		Me.aliveLed.CausesValidation = False
		Me.aliveLed.ledBackTransparent = False
		Me.aliveLed.ledBeep = ledAudio.audioOff
		Me.aliveLed.ledBlinkSpeed = blinkSpeed.blinkVeryFast
		Me.aliveLed.ledBorder = ledBorderStyle.None
		Me.aliveLed.ledColor = mainColor.colorGreen
		Me.aliveLed.ledColorBlink = blinkColor.colorDark
		Me.aliveLed.ledDesignBehaviour = ledDesignMode.blinkOn_audioOff
		Me.aliveLed.ledFlashNext = Nothing
		Me.aliveLed.ledNextTrigger = ledTrigger.nextDirect
		Me.aliveLed.ledSize = ledModel.sizeLarge
		Me.aliveLed.ledSoundFile = ""
		Me.aliveLed.ledStatus = ledModus.ledOn
		Me.aliveLed.Location = New System.Drawing.Point(480, 186)
		Me.aliveLed.Name = "aliveLed"
		Me.aliveLed.Size = New System.Drawing.Size(32, 32)
		Me.aliveLed.socketColor = System.Drawing.KnownColor.Silver
		Me.aliveLed.socketFeature = ledSocketFeature.socketRaised
		Me.aliveLed.socketWidth = ledSocketWidth.widthSmall
		Me.aliveLed.TabIndex = 17
		Me.aliveLed.toolTip = Nothing
		'
		'NotifyIcon
		'
		Me.NotifyIcon.Text = ""
		Me.NotifyIcon.Visible = True
		'
		'frmMitaSRV
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ClientSize = New System.Drawing.Size(524, 221)
		Me.Controls.Add(Me.aliveLed)
		Me.Controls.Add(Me.processLed)
		Me.Controls.Add(Me.Panel2)
		Me.Controls.Add(Me.SAPMld)
		Me.Controls.Add(Me.Label2)
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Location = New System.Drawing.Point(153, 425)
		Me.MaximizeBox = False
		Me.Name = "frmMitaSRV"
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Panel2.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub
#End Region

	Public WithEvents pscOrder As PSCMitaOrder.CMitaOrder
	Public WithEvents pscError As PscMitaError.CMitaError
	Public WithEvents pscConnect As pscMitaConnect.CMitaConnect
	Private mMenuItems(2) As MenuItem
	Private Sub initializeNotifyIcon()
		If Me.Visible Then
			mMenuItems(0) = New MenuItem("Hide " & startupInfo.myID & " Window", New EventHandler(AddressOf Me.hideMe))
		Else
			mMenuItems(0) = New MenuItem("Show " & startupInfo.myID & " Window", New EventHandler(AddressOf Me.showMe))
		End If
		mMenuItems(0).DefaultItem = True
		mMenuItems(1) = New MenuItem("-")
		mMenuItems(2) = New MenuItem("Shutdown " & startupInfo.myID, New EventHandler(AddressOf Me.endMe))
		Dim notifyiconMnu As ContextMenu = New ContextMenu(mMenuItems)
		NotifyIcon.ContextMenu = notifyiconMnu
	End Sub
	Private Sub frmMitaSRV_Closing(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		eventArgs.Cancel = Not EndIt
		EndIt = True
	End Sub

	Private Sub frmMitaSRV_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		Dim i As Short
		If Me.WindowState <> FormWindowState.Minimized Then
			With SAPMld
				.Width = Me.ClientRectangle.Width - 2 * .Left
				.Height = Me.ClientRectangle.Height - 3 * .Left - .Top
			End With
			aliveLed.Left = SAPMld.Left + SAPMld.Width - aliveLed.Width
			processLed.Left = aliveLed.Left - aliveLed.Width
			Label2.Top = Me.Height - Label2.Height - 30
			aliveLed.Top = Label2.Top - 12
			processLed.Top = Label2.Top - 12
			Label2.Width = SAPMld.Width - 3 * aliveLed.Width
		Else
			If Not Me.ShowInTaskbar And startupInfo.iconTray Then hideMe(Nothing, Nothing)
		End If
	End Sub

	Private Sub pscOrder_logContent(ByRef logText As String, ByVal logTyp As String) Handles pscOrder.logContent
		Select Case logTyp
			Case "L"
				listLogs.addItem(logText)
			Case "E"
				listErrors.addItem(logText)
			Case "W"
				listWarnings.addItem(logText)
		End Select
		listAll.addItem(logText)
		If (actLogTyp = logTyp Or actLogTyp = "A") And Not listStopped Then
			deleteList(SAPMld)
			SAPMld.Items.Insert(0, logText)
		End If
	End Sub

	Private Sub frmMitaSRV_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		applicationForm = Me
		processLed.ledBackTransparent = True
		aliveLed.ledBackTransparent = True
		init()
		pscConnect = mitaConnect
		pscOrder = SapOrder
		pscError = SapError
		Dim tmp As Bitmap = ImageList1.Images(startupInfo.iconIndex)
		Me.Icon = Icon.FromHandle(tmp.GetHicon)
		startupInfo.taskBar = True
		Me.Visible = False
		Dim id As Integer = 0
		Formloaded = True
		mitaData.applicationForm = Me
		mitaShared.connectSet = mitaConnect
		mitaShared.dataSet = mitaData
		mitaShared.systemSet = mitaSystem
		showApplication()
		doStartup(Me, SapOrder)
		mitaShared.appName = mitaApplication & "\" & CStr(myInstance)
		mitaData.caption = Me.Text
		mitaShared.RestPos(Me, True, True)
		If startupInfo.centerMe Then
			centerMe()
		ElseIf startupInfo.hideMe Then
			hideMe(Nothing, Nothing)
		Else
			showMe(Nothing, Nothing)
		End If
		Me.ShowInTaskbar = startupInfo.taskBar
		If startupInfo.iconTray Then
			NotifyIcon.Icon = Me.Icon
			NotifyIcon.Text = Me.Text
			NotifyIcon.Visible = True
			InitializeNotifyIcon()
		End If
		mitaData.processID = System.Diagnostics.Process.GetCurrentProcess.Id
		SapOrder.eventRaise("", mitaEventCodes.programStart)
		ToolBar1.Buttons.Item(1).Pushed = True
		ToolBarButton13.ToolTipText = "Shutdown " & mitaData.mitaApplication
		showLog(1)
		Timer1.Enabled = True
		'Timer2.Enabled = True
		If Not startupInfo.hideMe Then Me.Visible = True
		Me.Activate()
		aliveConnection.dbConnectString = connStr
		aliveQuery = "UPDATE " & mitaSystem.tableOnline & " SET alive = sysdate"
		aliveQuery = aliveQuery & " WHERE sapsystemid = " & startupInfo.sapSystemID
		aliveQuery = aliveQuery & " AND loginid = '" & CStr(myInstance) & "'"
		aliveQuery = aliveQuery & " AND loginhost = '" & startupInfo.hostName & "'"
		aliveQuery = aliveQuery & " AND loginapp = '" & startupInfo.myID & "'"
		aliveConnection.dataSet = mitaData
		SapOrder.init()
		workLoop()
	End Sub
	Private Sub workLoop()
		Dim a As String
		Dim path As String
		Dim i As Integer
		Dim idleCount As Integer = 0
		colorIdle()
		If Not EndIt Then
			Do
				System.Windows.Forms.Application.DoEvents()
				isIdle = False
				Do
					If processStopped Then
						Timer1.Enabled = True
						Exit Sub
					End If
					System.Windows.Forms.Application.DoEvents()
					abortOrder = False
					Select Case startupInfo.workModus
						Case workModi.inputDB
							If Not dbProcess() Then Exit Do
						Case workModi.inputDirectory
							If Not dirProcess() Then Exit Do
						Case workModi.inputSap
							If Not sapProcess() Then Exit Do
					End Select
					If EndIt Then Exit Do
					idleCount = 0
				Loop
				isIdle = True
				If EndIt Then Exit Do
				If doProfile Then
					idleCount = idleCount + 1
					If idleCount = 10 Then
						myProfile.profileClose()
						myProfile.profileOpen(Application.StartupPath & "\ProfilOutput\" & startupInfo.myID)
						idleCount = 0
					End If
				End If
				deleteList(SAPMld)
				Pause(2)
			Loop
		End If
		If doProfile Then myProfile.profileClose()
		SapOrder.eventRaise("", mitaEventCodes.programEnd)
		mitaShared.SavePos(Me)
		Me.Close()
	End Sub

	Private Sub deleteList(ByVal lst As Windows.Forms.ListBox)
		While lst.Items.Count > startupInfo.maxList
			lst.Items.Remove(lst.Items(lst.Items.Count - 1))
		End While
	End Sub

	Private Sub pscOrder_abortOrder() Handles pscOrder.abortOrder
		abortOrder = True
	End Sub

	Private Sub showLog(ByVal index As Integer)
		Dim i As Integer
		Dim anyPushed As Boolean
		If ToolBar1.Buttons.Item(index).Pushed Then
			For i = 1 To 4
				If i <> index Then ToolBar1.Buttons.Item(i).Pushed = False
			Next
		Else
			For i = 1 To 4
				anyPushed = anyPushed Or ToolBar1.Buttons.Item(i).Pushed
			Next
			If Not anyPushed Then ToolBar1.Buttons.Item(index).Pushed = True
			Exit Sub
		End If
		Select Case index
			Case 2
				actLogTyp = "L"
				fillListBox(SAPMld, listLogs)
			Case 4
				actLogTyp = "E"
				fillListBox(SAPMld, listErrors)
			Case 3
				actLogTyp = "W"
				fillListBox(SAPMld, listWarnings)
			Case 1
				actLogTyp = "A"
				fillListBox(SAPMld, listAll)
		End Select
	End Sub

	Private Sub optLog_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
		'showLog()
	End Sub

	Private Sub pscOrder_sapErrorForSend(ByRef sendDirect As Boolean) Handles pscOrder.sapErrorForSend
		sendDirect = Not startupInfo.batchError
	End Sub

	Private Sub pscError_sapError(ByRef message As String, ByVal code As Integer) Handles pscError.sapError
		pscOrder.eventRaise(message, code)
	End Sub

	Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
		Timer2.Enabled = False
		workLoop()
	End Sub

	Private Sub pscOrder_endApplication(ByRef immediatelly As Boolean) Handles pscOrder.endApplication
		'showApplication()
		If immediatelly Then
			Me.Close()
			End
		Else
			EndIt = True
			'showApplication()
		End If
	End Sub

	'Private Sub pscOrder_orderArrived() Handles pscOrder.orderArrived
	'	colorProcess()
	'	sapToDB()
	'	colorRestore()
	'End Sub

	Public Sub ToolBar1_ButtonClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolBarButtonClickEventArgs) Handles ToolBar1.ButtonClick
		Select Case ToolBar1.Buttons.IndexOf(e.Button)
			Case 1, 2, 3, 4
				showLog(ToolBar1.Buttons.IndexOf(e.Button))
			Case 0, 5, 7, 9, 11, 13, 15			 ' separator
			Case 6			 ' stop listing
				listStopped = ToolBar1.Buttons.Item(6).Pushed
				If listStopped Then
					ToolBar1.Buttons.Item(6).ImageIndex = 7
					ToolBar1.Buttons.Item(6).ToolTipText = "Continue List Refresh"
				Else
					ToolBar1.Buttons.Item(6).ImageIndex = 8
					ToolBar1.Buttons.Item(6).ToolTipText = "Stop List Refresh"
					Select Case actLogTyp
						Case "L"
							fillListBox(SAPMld, listLogs)
						Case "E"
							fillListBox(SAPMld, listErrors)
						Case "W"
							fillListBox(SAPMld, listWarnings)
						Case "A"
							fillListBox(SAPMld, listAll)
					End Select
				End If
			Case 8			 ' file
				listBoxToFile(SAPMld, "LogOutput", actLogTyp, True)
			Case 10			 ' stop process
				processStopped = ToolBar1.Buttons.Item(10).Pushed
				If processStopped Then
					ToolBar1.Buttons.Item(10).ImageIndex = 4
					ToolBar1.Buttons.Item(10).ToolTipText = "Continue Processing"
					colorStop()
				Else
					ToolBar1.Buttons.Item(10).ImageIndex = 5
					ToolBar1.Buttons.Item(10).ToolTipText = "Stop Processing"
					Timer2.Enabled = True
					colorIdle()
				End If
			Case 12			 ' about
				Dim about As pscMitaAbout.CMitaAbout = New pscMitaAbout.CMitaAbout
				about.Icon = Me.Icon
				about.dataSet = mitaData
				about.systemSet = mitaSystem
				about.sharedSet = mitaShared
				about.ShowDialog()
				about.Dispose()
			Case 14			 ' end
				endMe(Nothing, Nothing)
		End Select
	End Sub
	Private Sub endMe(ByVal sender As Object, ByVal e As EventArgs)
		EndIt = True
		If processStopped Then
			processStopped = False
			Timer2.Enabled = True
		End If
	End Sub
	Private Sub showMe(ByVal sender As Object, ByVal e As EventArgs)
		Me.Visible = True
		Me.WindowState = FormWindowState.Normal
		If Me.Top < 0 Or Me.Left < 0 Then
			centerMe()
		End If
		If startupInfo.iconTray Then InitializeNotifyIcon()
	End Sub
	Private Sub hideMe(ByVal sender As Object, ByVal e As EventArgs)
		Me.Visible = False
		If startupInfo.iconTray Then InitializeNotifyIcon()
	End Sub
	Private Sub centerMe()
		Dim scL As Rectangle = New Rectangle
		Dim scr As Screen = Screen.PrimaryScreen
		scL = scr.Bounds
		Dim x As Integer = (scL.Width - Me.Width) / 2
		Dim y As Integer = (scL.Height - Me.Height) / 2
		Me.Left = x
		Me.Top = y
	End Sub
	Private Sub SAPMld_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles SAPMld.MouseMove
		Dim MousePositionInClientCoords As Point = Me.SAPMld.PointToClient(Me.MousePosition)
		Dim indexUnderTheMouse As Integer = Me.SAPMld.IndexFromPoint(MousePositionInClientCoords)
		If indexUnderTheMouse > -1 Then
			Dim s As String = Me.SAPMld.Items(indexUnderTheMouse).ToString
			Dim g As Graphics = Me.SAPMld.CreateGraphics
			If g.MeasureString(s, Me.SAPMld.Font).Width > Me.SAPMld.ClientRectangle.Width Then
				Me.ToolTip1.SetToolTip(Me.SAPMld, s)
			Else
				Me.ToolTip1.SetToolTip(Me.SAPMld, "")
			End If
			g.Dispose()
		End If
	End Sub

	Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
		aliveLed.ledSwitch()
		If (isIdle Or processStopped) Then
			If tryAlive(mitaConnect) Then
				If isIdle Then
					pscOrder.eventRaise("$DATE$ $TIME$ Idle state, Alive written to DB", mitaEventCodes.userMessage)
				Else
					pscOrder.eventRaise("$DATE$ $TIME$ Process stopped, Alive written to DB", mitaEventCodes.userMessage)
				End If
			End If
		End If
	End Sub

	Private Sub frmMitaSRV_VisibleChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.VisibleChanged
		Me.Visible = Me.Visible And Not EndIt
	End Sub

	Private Sub mitaConnect_sqlError(ByVal description As String, ByVal query As String, ByVal title As String) Handles pscConnect.sqlError, pscOrder.sqlError
		Dim txt As String = description
		If query <> "" Then txt = txt & vbCrLf & query
		mitaMessage.waitTime = 0
		mitaMessage.popupMessage(txt, title)
	End Sub
End Class