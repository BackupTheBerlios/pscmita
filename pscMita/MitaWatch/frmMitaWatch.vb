Imports pscMitaDef.CMitaDef
Imports System.Data.Odbc
Public Class frmMitaWatch
	Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

	Public Sub New()
		MyBase.New()

		'This call is required by the Windows Form Designer.
		InitializeComponent()

		'Add any initialization after the InitializeComponent() call

	End Sub

	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If Not (components Is Nothing) Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub

	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer

	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.  
	'Do not modify it using the code editor.
	Friend WithEvents Panel1 As System.Windows.Forms.Panel
	Friend WithEvents ledAlarm As pscLedEx.pscLed
	Friend WithEvents btnEnd As System.Windows.Forms.Button
	Friend WithEvents btnReset As System.Windows.Forms.Button
	Friend WithEvents cmdSetup As System.Windows.Forms.Button
	Friend WithEvents ledTest As pscLedEx.pscLed
	Friend WithEvents lblStatus As System.Windows.Forms.Label
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents lblAlarm As System.Windows.Forms.Label
	Friend WithEvents Label17 As System.Windows.Forms.Label
	Friend WithEvents lblAlive As System.Windows.Forms.Label
	Friend WithEvents Label15 As System.Windows.Forms.Label
	Friend WithEvents lblLastOrder As System.Windows.Forms.Label
	Friend WithEvents Label13 As System.Windows.Forms.Label
	Friend WithEvents lblActOrder As System.Windows.Forms.Label
	Friend WithEvents Label12 As System.Windows.Forms.Label
	Friend WithEvents lblLogin As System.Windows.Forms.Label
	Friend WithEvents Label10 As System.Windows.Forms.Label
	Friend WithEvents lblID As System.Windows.Forms.Label
	Friend WithEvents Label8 As System.Windows.Forms.Label
	Friend WithEvents LblApp As System.Windows.Forms.Label
	Friend WithEvents Label6 As System.Windows.Forms.Label
	Friend WithEvents lblLogout As System.Windows.Forms.Label
	Friend WithEvents Label4 As System.Windows.Forms.Label
	Friend WithEvents lblHost As System.Windows.Forms.Label
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Panel2 As System.Windows.Forms.Panel
	Friend WithEvents lblProcess As System.Windows.Forms.Label
	Friend WithEvents Label5 As System.Windows.Forms.Label
	Friend WithEvents lblCaption As System.Windows.Forms.Label
	Friend WithEvents Label9 As System.Windows.Forms.Label
	Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
	Friend WithEvents btnShutDown As System.Windows.Forms.Button
	Friend WithEvents btnKill As System.Windows.Forms.Button
	Friend WithEvents btnStart As System.Windows.Forms.Button
	Friend WithEvents lblCmd As System.Windows.Forms.Label
	Friend WithEvents Label7 As System.Windows.Forms.Label
	Friend WithEvents refreshTimer As System.Windows.Forms.Timer
	Friend WithEvents cbSystem As System.Windows.Forms.ComboBox
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMitaWatch))
		Me.Panel1 = New System.Windows.Forms.Panel
		Me.ledAlarm = New pscLedEx.pscLed
		Me.btnEnd = New System.Windows.Forms.Button
		Me.btnReset = New System.Windows.Forms.Button
		Me.cmdSetup = New System.Windows.Forms.Button
		Me.ledTest = New pscLedEx.pscLed
		Me.lblStatus = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.lblAlarm = New System.Windows.Forms.Label
		Me.Label17 = New System.Windows.Forms.Label
		Me.lblAlive = New System.Windows.Forms.Label
		Me.Label15 = New System.Windows.Forms.Label
		Me.lblLastOrder = New System.Windows.Forms.Label
		Me.Label13 = New System.Windows.Forms.Label
		Me.lblActOrder = New System.Windows.Forms.Label
		Me.Label12 = New System.Windows.Forms.Label
		Me.lblLogin = New System.Windows.Forms.Label
		Me.Label10 = New System.Windows.Forms.Label
		Me.lblID = New System.Windows.Forms.Label
		Me.Label8 = New System.Windows.Forms.Label
		Me.LblApp = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.lblLogout = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.lblHost = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Panel2 = New System.Windows.Forms.Panel
		Me.lblCmd = New System.Windows.Forms.Label
		Me.Label7 = New System.Windows.Forms.Label
		Me.lblCaption = New System.Windows.Forms.Label
		Me.Label9 = New System.Windows.Forms.Label
		Me.lblProcess = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
		Me.btnShutDown = New System.Windows.Forms.Button
		Me.btnKill = New System.Windows.Forms.Button
		Me.btnStart = New System.Windows.Forms.Button
		Me.refreshTimer = New System.Windows.Forms.Timer(Me.components)
		Me.cbSystem = New System.Windows.Forms.ComboBox
		Me.Panel2.SuspendLayout()
		Me.SuspendLayout()
		'
		'Panel1
		'
		Me.Panel1.AutoScroll = True
		Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Panel1.Location = New System.Drawing.Point(8, 8)
		Me.Panel1.Name = "Panel1"
		Me.Panel1.Size = New System.Drawing.Size(300, 366)
		Me.Panel1.TabIndex = 0
		'
		'ledAlarm
		'
		Me.ledAlarm.CausesValidation = False
		Me.ledAlarm.ledBackTransparent = False
		Me.ledAlarm.ledBeep = pscLedEx.pscLed.ledAudio.soundOn
		Me.ledAlarm.ledBlinkSpeed = pscLedEx.pscLed.blinkSpeed.blinkMedium
		Me.ledAlarm.ledBorder = pscLedEx.pscLed.ledBorderStyle.Socket
		Me.ledAlarm.ledColor = pscLedEx.pscLed.mainColor.colorGreen
		Me.ledAlarm.ledColorBlink = pscLedEx.pscLed.blinkColor.colorDark
		Me.ledAlarm.ledDesignBehaviour = pscLedEx.pscLed.ledDesignMode.blinkOn_audioOff
		Me.ledAlarm.ledFlashNext = Nothing
		Me.ledAlarm.ledNextTrigger = pscLedEx.pscLed.ledTrigger.nextDirect
		Me.ledAlarm.ledSize = pscLedEx.pscLed.ledModel.sizeMaxi
		Me.ledAlarm.ledSoundFile = "pscLedEx.Phone.wav"
		Me.ledAlarm.ledStatus = pscLedEx.pscLed.ledModus.ledOn
		Me.ledAlarm.Location = New System.Drawing.Point(354, 278)
		Me.ledAlarm.Name = "ledAlarm"
		Me.ledAlarm.Size = New System.Drawing.Size(128, 128)
		Me.ledAlarm.socketColor = System.Drawing.KnownColor.PaleGoldenrod
		Me.ledAlarm.socketFeature = pscLedEx.pscLed.ledSocketFeature.socketRaised
		Me.ledAlarm.socketWidth = pscLedEx.pscLed.ledSocketWidth.widthMedium
		Me.ledAlarm.TabIndex = 2
		Me.ledAlarm.toolTip = "About ..."
		'
		'btnEnd
		'
		Me.btnEnd.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.btnEnd.Location = New System.Drawing.Point(594, 386)
		Me.btnEnd.Name = "btnEnd"
		Me.btnEnd.Size = New System.Drawing.Size(86, 20)
		Me.btnEnd.TabIndex = 3
		Me.btnEnd.Text = "Exit"
		'
		'btnReset
		'
		Me.btnReset.Enabled = False
		Me.btnReset.Location = New System.Drawing.Point(240, 388)
		Me.btnReset.Name = "btnReset"
		Me.btnReset.Size = New System.Drawing.Size(68, 20)
		Me.btnReset.TabIndex = 4
		Me.btnReset.Text = "Reset"
		'
		'cmdSetup
		'
		Me.cmdSetup.Location = New System.Drawing.Point(594, 288)
		Me.cmdSetup.Name = "cmdSetup"
		Me.cmdSetup.Size = New System.Drawing.Size(86, 20)
		Me.cmdSetup.TabIndex = 5
		Me.cmdSetup.Text = "Setup ..."
		'
		'ledTest
		'
		Me.ledTest.CausesValidation = False
		Me.ledTest.ledBackTransparent = False
		Me.ledTest.ledBeep = pscLedEx.pscLed.ledAudio.audioOff
		Me.ledTest.ledBlinkSpeed = pscLedEx.pscLed.blinkSpeed.blinkMedium
		Me.ledTest.ledBorder = pscLedEx.pscLed.ledBorderStyle.Socket
		Me.ledTest.ledColor = pscLedEx.pscLed.mainColor.colorRed
		Me.ledTest.ledColorBlink = pscLedEx.pscLed.blinkColor.colorDark
		Me.ledTest.ledDesignBehaviour = pscLedEx.pscLed.ledDesignMode.blinkOn_audioOff
		Me.ledTest.ledFlashNext = Nothing
		Me.ledTest.ledNextTrigger = pscLedEx.pscLed.ledTrigger.nextDirect
		Me.ledTest.ledSize = pscLedEx.pscLed.ledModel.sizeMedium
		Me.ledTest.ledSoundFile = ""
		Me.ledTest.ledStatus = pscLedEx.pscLed.ledModus.ledOn
		Me.ledTest.Location = New System.Drawing.Point(488, 380)
		Me.ledTest.Name = "ledTest"
		Me.ledTest.Size = New System.Drawing.Size(24, 24)
		Me.ledTest.socketColor = System.Drawing.KnownColor.Silver
		Me.ledTest.socketFeature = pscLedEx.pscLed.ledSocketFeature.sockedDeepened
		Me.ledTest.socketWidth = pscLedEx.pscLed.ledSocketWidth.widthMedium
		Me.ledTest.TabIndex = 6
		Me.ledTest.toolTip = Nothing
		Me.ledTest.Visible = False
		'
		'lblStatus
		'
		Me.lblStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblStatus.Location = New System.Drawing.Point(124, 175)
		Me.lblStatus.Name = "lblStatus"
		Me.lblStatus.Size = New System.Drawing.Size(182, 18)
		Me.lblStatus.TabIndex = 19
		'
		'Label3
		'
		Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(15, 175)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(92, 18)
		Me.Label3.TabIndex = 18
		Me.Label3.Text = "Status"
		'
		'lblAlarm
		'
		Me.lblAlarm.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblAlarm.Location = New System.Drawing.Point(124, 156)
		Me.lblAlarm.Name = "lblAlarm"
		Me.lblAlarm.Size = New System.Drawing.Size(182, 18)
		Me.lblAlarm.TabIndex = 17
		'
		'Label17
		'
		Me.Label17.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label17.Location = New System.Drawing.Point(15, 156)
		Me.Label17.Name = "Label17"
		Me.Label17.Size = New System.Drawing.Size(92, 18)
		Me.Label17.TabIndex = 16
		Me.Label17.Text = "Message"
		'
		'lblAlive
		'
		Me.lblAlive.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblAlive.Location = New System.Drawing.Point(124, 137)
		Me.lblAlive.Name = "lblAlive"
		Me.lblAlive.Size = New System.Drawing.Size(182, 18)
		Me.lblAlive.TabIndex = 15
		'
		'Label15
		'
		Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label15.Location = New System.Drawing.Point(15, 137)
		Me.Label15.Name = "Label15"
		Me.Label15.Size = New System.Drawing.Size(75, 18)
		Me.Label15.TabIndex = 14
		Me.Label15.Text = "Alive Time"
		'
		'lblLastOrder
		'
		Me.lblLastOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLastOrder.Location = New System.Drawing.Point(124, 118)
		Me.lblLastOrder.Name = "lblLastOrder"
		Me.lblLastOrder.Size = New System.Drawing.Size(182, 18)
		Me.lblLastOrder.TabIndex = 13
		'
		'Label13
		'
		Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label13.Location = New System.Drawing.Point(15, 118)
		Me.Label13.Name = "Label13"
		Me.Label13.Size = New System.Drawing.Size(101, 18)
		Me.Label13.TabIndex = 12
		Me.Label13.Text = "Previous Order"
		'
		'lblActOrder
		'
		Me.lblActOrder.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblActOrder.Location = New System.Drawing.Point(124, 99)
		Me.lblActOrder.Name = "lblActOrder"
		Me.lblActOrder.Size = New System.Drawing.Size(182, 18)
		Me.lblActOrder.TabIndex = 11
		'
		'Label12
		'
		Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label12.Location = New System.Drawing.Point(15, 99)
		Me.Label12.Name = "Label12"
		Me.Label12.Size = New System.Drawing.Size(92, 18)
		Me.Label12.TabIndex = 10
		Me.Label12.Text = "Aktual Order"
		'
		'lblLogin
		'
		Me.lblLogin.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLogin.Location = New System.Drawing.Point(124, 61)
		Me.lblLogin.Name = "lblLogin"
		Me.lblLogin.Size = New System.Drawing.Size(182, 18)
		Me.lblLogin.TabIndex = 9
		'
		'Label10
		'
		Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label10.Location = New System.Drawing.Point(15, 61)
		Me.Label10.Name = "Label10"
		Me.Label10.Size = New System.Drawing.Size(92, 18)
		Me.Label10.TabIndex = 8
		Me.Label10.Text = "Login Time"
		'
		'lblID
		'
		Me.lblID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblID.Location = New System.Drawing.Point(124, 42)
		Me.lblID.Name = "lblID"
		Me.lblID.Size = New System.Drawing.Size(182, 18)
		Me.lblID.TabIndex = 7
		'
		'Label8
		'
		Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label8.Location = New System.Drawing.Point(15, 42)
		Me.Label8.Name = "Label8"
		Me.Label8.Size = New System.Drawing.Size(92, 18)
		Me.Label8.TabIndex = 6
		Me.Label8.Text = "ID"
		'
		'LblApp
		'
		Me.LblApp.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LblApp.Location = New System.Drawing.Point(124, 23)
		Me.LblApp.Name = "LblApp"
		Me.LblApp.Size = New System.Drawing.Size(182, 18)
		Me.LblApp.TabIndex = 5
		'
		'Label6
		'
		Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.Location = New System.Drawing.Point(15, 23)
		Me.Label6.Name = "Label6"
		Me.Label6.Size = New System.Drawing.Size(92, 18)
		Me.Label6.TabIndex = 4
		Me.Label6.Text = "Application"
		'
		'lblLogout
		'
		Me.lblLogout.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblLogout.Location = New System.Drawing.Point(124, 80)
		Me.lblLogout.Name = "lblLogout"
		Me.lblLogout.Size = New System.Drawing.Size(182, 18)
		Me.lblLogout.TabIndex = 3
		'
		'Label4
		'
		Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Location = New System.Drawing.Point(15, 80)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(92, 18)
		Me.Label4.TabIndex = 2
		Me.Label4.Text = "Logout Time"
		'
		'lblHost
		'
		Me.lblHost.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblHost.Location = New System.Drawing.Point(124, 4)
		Me.lblHost.Name = "lblHost"
		Me.lblHost.Size = New System.Drawing.Size(182, 18)
		Me.lblHost.TabIndex = 1
		'
		'Label1
		'
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(15, 4)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(92, 18)
		Me.Label1.TabIndex = 0
		Me.Label1.Text = "Host"
		'
		'Panel2
		'
		Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.Panel2.Controls.Add(Me.lblCmd)
		Me.Panel2.Controls.Add(Me.Label7)
		Me.Panel2.Controls.Add(Me.lblCaption)
		Me.Panel2.Controls.Add(Me.Label9)
		Me.Panel2.Controls.Add(Me.lblProcess)
		Me.Panel2.Controls.Add(Me.Label5)
		Me.Panel2.Controls.Add(Me.lblStatus)
		Me.Panel2.Controls.Add(Me.Label3)
		Me.Panel2.Controls.Add(Me.lblAlarm)
		Me.Panel2.Controls.Add(Me.Label17)
		Me.Panel2.Controls.Add(Me.lblAlive)
		Me.Panel2.Controls.Add(Me.Label15)
		Me.Panel2.Controls.Add(Me.lblLastOrder)
		Me.Panel2.Controls.Add(Me.Label13)
		Me.Panel2.Controls.Add(Me.lblActOrder)
		Me.Panel2.Controls.Add(Me.Label12)
		Me.Panel2.Controls.Add(Me.lblLogin)
		Me.Panel2.Controls.Add(Me.Label10)
		Me.Panel2.Controls.Add(Me.lblID)
		Me.Panel2.Controls.Add(Me.Label8)
		Me.Panel2.Controls.Add(Me.LblApp)
		Me.Panel2.Controls.Add(Me.Label6)
		Me.Panel2.Controls.Add(Me.lblLogout)
		Me.Panel2.Controls.Add(Me.Label4)
		Me.Panel2.Controls.Add(Me.lblHost)
		Me.Panel2.Controls.Add(Me.Label1)
		Me.Panel2.Location = New System.Drawing.Point(352, 6)
		Me.Panel2.Name = "Panel2"
		Me.Panel2.Size = New System.Drawing.Size(326, 264)
		Me.Panel2.TabIndex = 1
		'
		'lblCmd
		'
		Me.lblCmd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCmd.Location = New System.Drawing.Point(124, 232)
		Me.lblCmd.Name = "lblCmd"
		Me.lblCmd.Size = New System.Drawing.Size(182, 18)
		Me.lblCmd.TabIndex = 25
		'
		'Label7
		'
		Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.Location = New System.Drawing.Point(15, 232)
		Me.Label7.Name = "Label7"
		Me.Label7.Size = New System.Drawing.Size(101, 18)
		Me.Label7.TabIndex = 24
		Me.Label7.Text = "Arguments"
		'
		'lblCaption
		'
		Me.lblCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblCaption.Location = New System.Drawing.Point(124, 213)
		Me.lblCaption.Name = "lblCaption"
		Me.lblCaption.Size = New System.Drawing.Size(182, 18)
		Me.lblCaption.TabIndex = 23
		'
		'Label9
		'
		Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label9.Location = New System.Drawing.Point(15, 213)
		Me.Label9.Name = "Label9"
		Me.Label9.Size = New System.Drawing.Size(101, 18)
		Me.Label9.TabIndex = 22
		Me.Label9.Text = "Caption"
		'
		'lblProcess
		'
		Me.lblProcess.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblProcess.Location = New System.Drawing.Point(124, 194)
		Me.lblProcess.Name = "lblProcess"
		Me.lblProcess.Size = New System.Drawing.Size(182, 18)
		Me.lblProcess.TabIndex = 21
		'
		'Label5
		'
		Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.Location = New System.Drawing.Point(15, 194)
		Me.Label5.Name = "Label5"
		Me.Label5.Size = New System.Drawing.Size(79, 18)
		Me.Label5.TabIndex = 20
		Me.Label5.Text = "Process ID"
		'
		'btnShutDown
		'
		Me.btnShutDown.Enabled = False
		Me.btnShutDown.Location = New System.Drawing.Point(8, 388)
		Me.btnShutDown.Name = "btnShutDown"
		Me.btnShutDown.Size = New System.Drawing.Size(68, 20)
		Me.btnShutDown.TabIndex = 7
		Me.btnShutDown.Text = "ShutDown"
		'
		'btnKill
		'
		Me.btnKill.Enabled = False
		Me.btnKill.Location = New System.Drawing.Point(82, 388)
		Me.btnKill.Name = "btnKill"
		Me.btnKill.Size = New System.Drawing.Size(68, 20)
		Me.btnKill.TabIndex = 8
		Me.btnKill.Text = "Kill"
		'
		'btnStart
		'
		Me.btnStart.Location = New System.Drawing.Point(594, 314)
		Me.btnStart.Name = "btnStart"
		Me.btnStart.Size = New System.Drawing.Size(86, 20)
		Me.btnStart.TabIndex = 9
		Me.btnStart.Text = "Start ..."
		'
		'refreshTimer
		'
		Me.refreshTimer.Interval = 10000
		'
		'cbSystem
		'
		Me.cbSystem.Location = New System.Drawing.Point(496, 292)
		Me.cbSystem.Name = "cbSystem"
		Me.cbSystem.Size = New System.Drawing.Size(82, 21)
		Me.cbSystem.TabIndex = 10
		'
		'frmMitaWatch
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ClientSize = New System.Drawing.Size(690, 417)
		Me.Controls.Add(Me.cbSystem)
		Me.Controls.Add(Me.btnStart)
		Me.Controls.Add(Me.btnKill)
		Me.Controls.Add(Me.btnShutDown)
		Me.Controls.Add(Me.ledTest)
		Me.Controls.Add(Me.cmdSetup)
		Me.Controls.Add(Me.btnReset)
		Me.Controls.Add(Me.btnEnd)
		Me.Controls.Add(Me.ledAlarm)
		Me.Controls.Add(Me.Panel2)
		Me.Controls.Add(Me.Panel1)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.Name = "frmMitaWatch"
		Me.Text = "pscMitaWatch"
		Me.Panel2.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub

#End Region
	Friend mitaApplication As String
	Public watchTarget() As pscMitaCand.pscCand
	Dim actTarget As pscMitaCand.pscCand
	Dim testTarget As pscMitaCand.pscCand
	Dim sampleInterval As Integer
	Dim timeOut As Integer
	Dim audio As pscLedEx.pscLed.ledAudio
	Dim sound As String
	Dim agent As clsMitaProcess
	Dim isStarting As Boolean
	Dim hostList() As String
	Dim progList() As String
	Dim candTop As Integer = 5
	Dim targetCount As Integer = -1
	Private Sub frmMitaWatch_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		Dim user As String
		Dim db As String
		Dim pwd As String
		isStarting = True
		mainFrm = Me
		'Dim hostName As String = System.Net.Dns.GetHostName
		'Dim host As System.Net.IPHostEntry = System.Net.Dns.GetHostByName(hostName)
		'Dim addresses As System.Net.IPAddress() = host.AddressList
		'Dim b As String = addresses(0).ToString()
		'Dim c As String = addresses(1).ToString()
		'Dim d As String = addresses(2).ToString()
		mitaApplication = "PSC\MitaWatch"
		mitaData.registryApplication = mitaApplication
		mitaShared.appName = mitaApplication
		mitaShared.RestPos(Me, False)
		mitaShared.systemSet = mitaSystem
		mitaShared.connectSet = mitaConnect
		mitaShared.dataSet = mitaData
		Me.Show()
		Application.DoEvents()
		timeOut = CInt(GetSetting(mitaApplication, "Settings", "TimeOut", "120"))
		sampleInterval = CInt(GetSetting(mitaApplication, "Settings", "SampleInterval", "10"))
		sound = GetSetting(mitaApplication, "Settings", "Sound", "")
		audio = CType(GetSetting(mitaApplication, "Settings", "Audio", "0"), pscLedEx.pscLed.ledAudio)
		Dim loginCls As New pscMitaLogin.CLogin
		loginCls.registryKey = mitaApplication
		loginCls.popUp = True
		If Not loginCls.doLogin Then End
		user = loginCls.user
		db = loginCls.dataBase
		mitaConnect.dbConnectString = loginCls.connectString
		loginCls.Dispose()
		readSap()
		addToList(hostList, System.Net.Dns.GetHostName)
		If IsNothing(agent) Then agent = New clsMitaProcess
		agent.caption = Nothing
		agent.hostName = Nothing
		setTimes()
		ledAlarm.ledSoundFile = sound
		ledAlarm.ledBeep = audio
		isStarting = False
		If Not IsNothing(watchTarget) Then
			watchTarget(0).Selected = True
			refreshTimer.Enabled = True
		End If
	End Sub
	Public Sub doTheQuery(ByVal rePos As Boolean)
		Dim query As String
		Dim i As Integer
		candTop = 5
		Panel1.Size = New Size(300, Panel1.Height)
		If Not IsNothing(watchTarget) And Not rePos Then
			For i = 0 To UBound(watchTarget)
				watchTarget(i).Dispose()
			Next
			progList = Nothing
			watchTarget = Nothing
			targetCount = -1
		End If

		query = "SELECT loginhost, loginid, loginapp FROM " & mitaSystem.tableOnline
		query = query & " ORDER BY loginid ASC"
		mitaConnect.odbc_connection.ConnectionString = mitaConnect.dbConnectString
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		idbc.CommandText = query
		Try
			mitaConnect.odbc_connection.Open()
			reader = idbc.ExecuteReader()
			While reader.Read
				Me.SuspendLayout()
				targetCount = targetCount + 1
				ReDim Preserve watchTarget(targetCount)
				watchTarget(targetCount) = New pscMitaCand.pscCand
				testTarget = watchTarget(targetCount)
				testTarget.Name = "PscCand" & CStr(targetCount)
				If addCandidate(testTarget, candTop, reader) Then
					actTarget = watchTarget(targetCount)
					If rePos Then
						posTarget(candTop, actTarget.sortNumber)
						actTarget.Selected = True
					End If
				Else
					If rePos Then posTarget(candTop, testTarget.sortNumber)
					testTarget.Dispose()
					targetCount = targetCount - 1
				End If
				candTop = candTop + 30
				Me.ResumeLayout()
			End While
			reader.Close()
			idbc.Dispose()
		Catch
			MsgBox(Err.Description & vbCrLf & query)
		End Try
		mitaConnect.odbc_connection.Close()
		If candTop > Panel1.Height Then
			Panel1.AutoScrollMinSize = New Size(0, candTop)
			Panel1.Size = New Size(Panel1.Width + 15, Panel1.Height)
			Panel1.Refresh()
		End If
	End Sub
	Private Sub posTarget(ByVal myTop As Integer, ByVal key As Integer)
		Dim i As Integer
		For i = 0 To targetCount
			If watchTarget(i).sortNumber = key Then
				watchTarget(i).Top = myTop
				Exit Sub
			End If
		Next
		i = 0
	End Sub
	Private Function addToList(ByRef list() As String, ByVal item As String) As Boolean
		Dim i As Integer = 0
		If IsNothing(list) Then
			ReDim list(i)
			list(i) = item
			Return True
		End If
		For i = 0 To UBound(list)
			If list(i) = item Then Return False
		Next
		ReDim Preserve list(i)
		list(i) = item
		Return True
	End Function
	Private Function addCandidate(ByVal target As pscMitaCand.pscCand, ByVal myTop As Integer, ByVal reader As OdbcDataReader) As Boolean
		Dim isNew As Boolean
		Dim id As Integer
		AddHandler target.dbContent, AddressOf watchTarget_dbContent
		AddHandler target.mitaAlarm, AddressOf watchTarget_mitaAlarm
		target.connectString = mitaConnect.dbConnectString
		target.programName = reader.Item("loginapp").ToString
		id = CInt(reader.Item("loginid").ToString)
		target.hostID = id
		target.hostName = reader.Item("loginhost").ToString
		target.onlineTable = mitaSystem.tableOnline
		addToList(hostList, target.hostName)
		isNew = addToList(progList, target.programName & CStr(id))
		If isNew Then
			target.Location = New System.Drawing.Point(5, myTop)
			target.Size = New System.Drawing.Size(232, 28)
			Me.Panel1.Controls.Add(target)
			target.startUp()
		End If
		Return isNew
	End Function
	Private Sub watchTarget_mitaAlarm(ByVal sender As System.Windows.Forms.Control, ByVal message As String)
		actTarget = CType(sender, pscMitaCand.pscCand)
		ledAlarm.ledStatus = pscLedEx.pscLed.ledModus.ledBlink
		ledAlarm.ledColor = pscLedEx.pscLed.mainColor.colorRed
		Panel1.ScrollControlIntoView(actTarget)
	End Sub
	Private Sub watchTarget_dbContent(ByVal sender As System.Windows.Forms.Control, ByVal dbLine As pscMitaCand.pscCand.dbData)
		actTarget = CType(sender, pscMitaCand.pscCand)
		btnReset.Enabled = (actTarget.Status = pscMitaCand.pscCand.mitaStatus.Problem) Or (actTarget.Status = pscMitaCand.pscCand.mitaStatus.TimedOut)
		btnKill.Enabled = btnReset.Enabled
		btnShutDown.Enabled = (actTarget.Status = pscMitaCand.pscCand.mitaStatus.Running) Or btnReset.Enabled
		Me.lblHost.Text = dbLine.loginHost
		Me.LblApp.Text = dbLine.loginApp
		Me.lblID.Text = dbLine.loginId
		Me.lblActOrder.Text = dbLine.orderNo
		Me.lblLastOrder.Text = dbLine.lastOrder
		Me.lblLogin.Text = dbLine.loginTime
		Me.lblLogout.Text = dbLine.logoutTime
		Me.lblAlive.Text = dbLine.alive
		Me.lblAlarm.Text = dbLine.alarm
		Me.lblProcess.Text = dbLine.processId
		Me.lblCaption.Text = dbLine.caption
		Me.ToolTip1.SetToolTip(lblCaption, lblCaption.Text)
		Me.lblCmd.Text = dbLine.command
		Me.ToolTip1.SetToolTip(lblCmd, lblCmd.Text)
		Dim tmp As String
		Select Case actTarget.Status
			Case pscMitaCand.pscCand.mitaStatus.Running
				tmp = "Online"
			Case pscMitaCand.pscCand.mitaStatus.NotRunning
				tmp = "Offline"
			Case pscMitaCand.pscCand.mitaStatus.Waiting
				tmp = "Waiting for First Alive"
			Case pscMitaCand.pscCand.mitaStatus.Problem
				tmp = "Problem"
			Case pscMitaCand.pscCand.mitaStatus.Starting
				tmp = "Unknown"
			Case pscMitaCand.pscCand.mitaStatus.TimedOut
				tmp = "Timed Out"
		End Select
		Me.lblStatus.Text = tmp
		startAgent()
	End Sub
	Private Sub startAgent()
		If IsNothing(agent) Then agent = New clsMitaProcess
		agent.caption = lblCaption.Text
		If lblCaption.Text = "" Then agent.caption = Nothing
		agent.hostName = lblHost.Text
		If lblHost.Text = "" Then agent.hostName = Nothing
	End Sub
	Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
		actTarget.startUp()
		Dim i As Integer
		For i = 0 To UBound(watchTarget)
			If watchTarget(i).Status = pscMitaCand.pscCand.mitaStatus.TimedOut _
			Or watchTarget(i).Status = pscMitaCand.pscCand.mitaStatus.Problem Then
				watchTarget(i).Selected = True
				Exit Sub
			End If
		Next
		ledAlarm.ledColor = pscLedEx.pscLed.mainColor.colorGreen
		ledAlarm.ledStatus = pscLedEx.pscLed.ledModus.ledOn
	End Sub

	Private Sub btnEnd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnd.Click
		Dim i As Integer
		If Not IsNothing(watchTarget) Then
			For i = 0 To UBound(watchTarget)
				If Not IsNothing(watchTarget(i)) Then watchTarget(i).Dispose()
			Next
		End If
		ledAlarm.Dispose()
		End
	End Sub

	Private Sub cmdSetup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSetup.Click
		Dim i As Integer
		Dim frm As New frmSetup
		Dim res As System.Windows.Forms.DialogResult
		frm.Led = ledAlarm
		frm.Timeout = timeOut
		frm.SampleInterval = sampleInterval
		frm.Sound = sound
		frm.Audio = audio
		res = frm.ShowDialog()
		If res = DialogResult.Cancel Then
			res = res
		Else
			timeOut = frm.Timeout
			sampleInterval = frm.SampleInterval
			sound = frm.Sound
			audio = frm.Audio
			setTimes()
			ledAlarm.ledSoundFile = sound
			ledAlarm.ledBeep = audio
			If audio <> pscLedEx.pscLed.ledAudio.audioOff Then
				ledTest.ledSoundFile = sound
				ledTest.ledBeep = audio
				ledTest.ledStatus = pscLedEx.pscLed.ledModus.ledFlash
			End If
			SaveSetting(mitaApplication, "Settings", "TimeOut", CStr(timeOut))
			SaveSetting(mitaApplication, "Settings", "SampleInterval", CInt(sampleInterval))
			SaveSetting(mitaApplication, "Settings", "Sound", sound)
			SaveSetting(mitaApplication, "Settings", "Audio", CStr(audio))
		End If
		frm.Dispose()
	End Sub
	Private Sub setTimes()
		Dim i As Integer
		If IsNothing(watchTarget) Then Exit Sub
		For i = 0 To UBound(watchTarget)
			watchTarget(i).TimeOut = timeOut
			watchTarget(i).sampleInterval = sampleInterval
		Next
	End Sub

	Private Sub frmMitaWatch_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		Dim i As Integer
		mitaShared.SavePos(Me)
		For i = 0 To UBound(watchTarget)
			watchTarget(i).Dispose()
		Next i
		ledAlarm.Dispose()
		ledTest.Dispose()
	End Sub

	Private Sub btnShutDown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShutDown.Click
		Dim result As Boolean
		If agent.shutdownMitaApplication() Then
			actTarget.actualize()
		Else
			MsgBox("ShutDown Successless")
			btnKill.Enabled = True
		End If
	End Sub

	Private Sub btnKill_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnKill.Click
		If agent.killMitaApplication() Then
			MsgBox("Kill Succeeded")
			actTarget.isKilled()
		Else
			MsgBox("Kill Successless")
		End If
	End Sub

	Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
		Dim result As DialogResult
		Dim start As New frmStart
		Dim systems() As String
		Dim indexes() As Integer
		Dim count As Integer
		Dim i As Integer
		start.hosts = hostList
		mitaShared.readDBSap(mitaConnect.odbc_connection, systems, indexes, count)
		start.sapsystems = systems
		start.bldProgs()
		If Not IsNothing(actTarget) Then
			start.host = actTarget.hostName
			start.program = actTarget.programName
			start.arguments = actTarget.arguments
		End If
		result = start.ShowDialog
		If result = DialogResult.OK Then
			agent.hostName = start.host
			If agent.startMitaApplication(start.program, start.arguments) Then
				refreshTimer.Enabled = False
				refreshTimer.Interval = 3000
				refreshTimer.Enabled = True
			Else
				MsgBox("Could not start")
			End If
		End If
		start.Dispose()
	End Sub

	Private Sub refreshTimer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles refreshTimer.Tick
		Dim query As String
		refreshTimer.Enabled = False
		doTheQuery(True)
		setTimes()
		refreshTimer.Enabled = True
		actTarget.Selected = True
	End Sub

	Private Sub ledAlarm_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ledAlarm.Click
		mitaData.createHost = System.Net.Dns.GetHostName
		mitaShared.dataSet = mitaData
		Dim about As pscMitaAbout.CMitaAbout = New pscMitaAbout.CMitaAbout
		about.Icon = Me.Icon
		about.sharedSet = mitaShared
		about.ShowDialog()
		about.Dispose()
	End Sub

	Private Sub cbSystem_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSystem.SelectedIndexChanged
		If cbSystem.SelectedIndex = -1 Then Exit Sub
		If cbSystem.Text = "" Then Exit Sub
		refreshTimer.Enabled = False
		actSap = cbSystem.Text
		mitaShared.readDBSapSystemFromName(actSap)
		mitaConnect.dbConnectString = mitaSystem.connectString
		mitaShared.buildVersionInfo(Me)
		refreshTimer.Enabled = True
		doTheQuery(False)
	End Sub
End Class
