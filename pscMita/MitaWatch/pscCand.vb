Imports pscMitaDef.CMitaDef
Imports System.Data
Imports System.Data.Odbc
Imports System.Windows.Forms
Imports pscLedEx.pscLed

Public Class pscCand
	Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

	Public Sub New()
		MyBase.New()

		'This call is required by the Windows Form Designer.
		InitializeComponent()

		'Add any initialization after the InitializeComponent() call

	End Sub

	'UserControl overrides dispose to clean up the component list.
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
	Friend WithEvents Timer1 As System.Windows.Forms.Timer
	Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
	Friend WithEvents lblID As System.Windows.Forms.Label
	Friend WithEvents lblHost As System.Windows.Forms.Label
	Friend WithEvents lblProgram As System.Windows.Forms.Label
	Friend WithEvents Timer2 As System.Windows.Forms.Timer
	Friend WithEvents Led As pscLedEx.pscLed
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
		Me.PictureBox1 = New System.Windows.Forms.PictureBox
		Me.lblID = New System.Windows.Forms.Label
		Me.lblHost = New System.Windows.Forms.Label
		Me.lblProgram = New System.Windows.Forms.Label
		Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
		Me.Led = New pscLedEx.pscLed
		Me.SuspendLayout()
		'
		'Timer1
		'
		Me.Timer1.Interval = 19000
		'
		'PictureBox1
		'
		Me.PictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
		Me.PictureBox1.Name = "PictureBox1"
		Me.PictureBox1.Size = New System.Drawing.Size(284, 28)
		Me.PictureBox1.TabIndex = 4
		Me.PictureBox1.TabStop = False
		'
		'lblID
		'
		Me.lblID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblID.Location = New System.Drawing.Point(146, 2)
		Me.lblID.Name = "lblID"
		Me.lblID.Size = New System.Drawing.Size(40, 24)
		Me.lblID.TabIndex = 8
		Me.lblID.Text = "100"
		Me.lblID.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		'
		'lblHost
		'
		Me.lblHost.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblHost.Location = New System.Drawing.Point(186, 2)
		Me.lblHost.Name = "lblHost"
		Me.lblHost.Size = New System.Drawing.Size(93, 24)
		Me.lblHost.TabIndex = 7
		Me.lblHost.Text = "Label1"
		Me.lblHost.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		'
		'lblProgram
		'
		Me.lblProgram.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblProgram.Location = New System.Drawing.Point(6, 2)
		Me.lblProgram.Name = "lblProgram"
		Me.lblProgram.Size = New System.Drawing.Size(140, 24)
		Me.lblProgram.TabIndex = 6
		Me.lblProgram.Text = "Label1"
		Me.lblProgram.TextAlign = System.Drawing.ContentAlignment.MiddleRight
		'
		'Timer2
		'
		Me.Timer2.Interval = 1000
		'
		'Led
		'
		Me.Led.CausesValidation = False
		Me.Led.ledBeep = pscLedEx.pscLed.ledAudio.audioOff
		Me.Led.ledBlinkSpeed = pscLedEx.pscLed.blinkSpeed.blinkMedium
		Me.Led.ledBorder = pscLedEx.pscLed.ledBorderStyle.Socket
		Me.Led.ledColor = pscLedEx.pscLed.mainColor.colorYellow
		Me.Led.ledColorBlink = pscLedEx.pscLed.blinkColor.colorDark
		Me.Led.ledDesignBehaviour = pscLedEx.pscLed.ledDesignMode.blinkOn_audioOff
		Me.Led.ledFlashNext = Nothing
		Me.Led.ledNextTrigger = pscLedEx.pscLed.ledTrigger.nextDirect
		Me.Led.ledSize = pscLedEx.pscLed.ledModel.sizeMedium
		Me.Led.ledSoundFile = ""
		Me.Led.ledStatus = pscLedEx.pscLed.ledModus.ledOn
		Me.Led.Location = New System.Drawing.Point(2, 2)
		Me.Led.Name = "Led"
		Me.Led.Size = New System.Drawing.Size(24, 24)
		Me.Led.socketFeature = pscLedEx.pscLed.ledSocketFeature.sockedFlat
		Me.Led.socketWidth = pscLedEx.pscLed.ledSocketWidth.widthMedium
		Me.Led.TabIndex = 9
		Me.Led.toolTip = Nothing
		'
		'pscCand
		'
		Me.Controls.Add(Me.Led)
		Me.Controls.Add(Me.lblID)
		Me.Controls.Add(Me.lblHost)
		Me.Controls.Add(Me.lblProgram)
		Me.Controls.Add(Me.PictureBox1)
		Me.Name = "pscCand"
		Me.Size = New System.Drawing.Size(286, 38)
		Me.ResumeLayout(False)

	End Sub

#End Region
#Region "Structures"
	Public Structure dbData
		Dim sapSystemID As Integer
		Dim loginTime As Date
		Dim logoutTime As Date
		Dim loginHost As String
		Dim loginId As String
		Dim loginApp As String
		Dim orderNo As String
		Dim lastOrder As String
		Dim alive As Date
		Dim alarm As String
		Dim processId As String
		Dim caption As String
		Dim command As String
	End Structure
	Public Enum mitaStatus
		Starting
		Waiting
		Running
		Problem
		TimedOut
		NotRunning
	End Enum
#End Region
#Region "Local Variables and Events"
	Private mvarHost As String = ""
	Private mvarID As Integer = -9999
	Private mvarProgram As String = ""
	Private mvarConnectString As String = ""
	Private mvarQuery As String = ""
	Private mvarTable As String = ""
	Private odbc_connection As New OdbcConnection
	Private mvarInterval As Integer = 20000
	Private mvarDbLine As dbData
	Private mvarLastAlive As Date = Nothing
	Private mvarStartAlive As Date = Nothing
	Private mvarAliveTime As Date
	Private mvarWaitStartTime As Date
	Private mvarLastAlarm As String = ""
	Private mvarIsAlarm As Boolean = False
	Private mvarStatus As mitaStatus
	Private mvarTimeOut As Integer = 20

	Private Const offlineColor = mainColor.colorCyan
	Private Const testColor = mainColor.colorYellow
	Private Const okColor = mainColor.colorGreen
	Private Const problemColor = mainColor.colorRed

	Public Event dbContent(ByVal sender As Control, ByVal dbLine As dbData)
	Public Event mitaAlarm(ByVal sender As Control, ByVal message As String)
#End Region
#Region "Overrides"
	Protected Overrides Sub OnResize(ByVal e As System.EventArgs)
		Me.Width = 286
		Me.Height = 28
	End Sub
#End Region
#Region "Properties"
	Property Selected() As Boolean
		Get
			Return Led.ledBorder <> ledBorderStyle.Socket
		End Get
		Set(ByVal Value As Boolean)
			If Value Then
				Led.socketFeature = ledSocketFeature.sockedFlat
				'Led.ledBorder = ledBorderStyle.None
				'Led.ledSize = ledModel.sizeMedium
				'Led.Top = 2
				'Led.Left = 2
				inactivateLeds()
				sendContent()
				lblProgram.ForeColor = System.Drawing.Color.Blue
				lblID.ForeColor = System.Drawing.Color.Blue
				lblHost.ForeColor = System.Drawing.Color.Blue
			Else
				Led.socketFeature = ledSocketFeature.sockedDeepened
				'Led.ledBorder = ledBorderStyle.Socket
				'Led.ledSize = ledModel.sizeSmall
				'Led.Top = 6
				'Led.Left = 6
				lblProgram.ForeColor = System.Drawing.Color.Black
				lblID.ForeColor = System.Drawing.Color.Black
				lblHost.ForeColor = System.Drawing.Color.Black
			End If
		End Set
	End Property
	Property TimeOut() As Integer
		Get
			Return mvarTimeOut
		End Get
		Set(ByVal Value As Integer)
			mvarTimeOut = Value
		End Set
	End Property
	WriteOnly Property sampleInterval() As Integer
		Set(ByVal Value As Integer)
			If Value < 2 Then
				mvarInterval = 2000
			Else
				mvarInterval = (Value - 1) * 1000
			End If
			Timer1.Interval = mvarInterval
		End Set
	End Property
	Property hostName() As String
		Get
			Return mvarHost
		End Get
		Set(ByVal Value As String)
			mvarHost = Value
			lblHost.Text = Value
		End Set
	End Property
	ReadOnly Property sap() As Integer
		Get
			Return mvarDbLine.sapSystemID
		End Get
	End Property
	ReadOnly Property arguments() As String
		Get
			Return mvarDbLine.command
		End Get
	End Property
	ReadOnly Property sortNumber() As Integer
		Get
			Return mvarID
		End Get
	End Property
	WriteOnly Property hostID() As Integer
		Set(ByVal Value As Integer)
			mvarID = Value
			lblID.Text = Value
		End Set
	End Property
	Property programName() As String
		Get
			Return mvarProgram
		End Get
		Set(ByVal Value As String)
			mvarProgram = Value
			lblProgram.Text = Value
		End Set
	End Property
	WriteOnly Property connectString() As String
		Set(ByVal Value As String)
			mvarConnectString = Value
		End Set
	End Property
	WriteOnly Property onlineTable() As String
		Set(ByVal Value As String)
			mvarTable = Value
		End Set
	End Property
	ReadOnly Property Status() As mitaStatus
		Get
			Return mvarStatus
		End Get
	End Property
#End Region
#Region "Public Functions"
	Public Function startUp() As Boolean
		If mvarID <> -9999 And mvarHost <> "" And mvarProgram <> "" And mvarConnectString <> "" And mvarTable <> "" Then
			If connectionOpenOdbc() Then
				cleanHistory()
				Timer1.Interval = mvarInterval
				Timer2_Tick(Nothing, Nothing)
				Led.ledColor = testColor
				sendContent()
				Return True
			End If
		End If
		Return False
	End Function
	Public Sub shutOff()
		Timer1.Enabled = False
	End Sub
	Public Sub isKilled()
		Dim query As String
		query = "UPDATE " & mvarTable
		query = query & " SET logouttime = sysdate, alive = NULL, alarm = 'KILLED by operator', caption = NULL, processid = NULL, command = NULL"
		query = query & " WHERE loginhost = '" & mvarHost & "'"
		query = query & " AND loginid = " & CInt(mvarID)
		query = query & " And loginapp = '" & mvarProgram & "'"
		Dim idbc As OdbcCommand = odbc_connection.CreateCommand()
		idbc.CommandText = query
		Try
			odbc_connection.Open()
			idbc.ExecuteNonQuery()
		Catch
			MsgBox(Err.Description & vbCrLf & mvarQuery)
		End Try
		idbc.Dispose()
		odbc_connection.Close()
		Timer1_Tick(Nothing, Nothing)
	End Sub
	Public Sub isOffline()
		Dim query As String
		query = "UPDATE " & mvarTable
		query = query & " SET logouttime = sysdate, alive = NULL, processid = NULL, caption = NULL"
		query = query & " WHERE loginhost = '" & mvarHost & "'"
		query = query & " AND loginid = " & CInt(mvarID)
		query = query & " And loginapp = '" & mvarProgram & "'"
		Dim idbc As OdbcCommand = odbc_connection.CreateCommand()
		idbc.CommandText = query
		Try
			odbc_connection.Open()
			idbc.ExecuteNonQuery()
		Catch
			MsgBox(Err.Description & vbCrLf & mvarQuery)
		End Try
		idbc.Dispose()
		odbc_connection.Close()
		Timer1_Tick(Nothing, Nothing)
	End Sub
	Public Sub isShutDown()
		Timer1_Tick(Nothing, Nothing)
	End Sub
	Public Sub isStarted()
		Timer1_Tick(Nothing, Nothing)
	End Sub
	Public Sub actualize()
		Timer2_Tick(Nothing, Nothing)
	End Sub
#End Region
#Region "Private Functions"
	Private Sub sendContent()
		RaiseEvent dbContent(Me, mvarDbLine)
	End Sub
	Private Sub inactivateLeds()
		Dim x As Control
		Dim tmp As pscCand
		For Each x In Me.Parent.Controls
			If TypeOf (x) Is pscCand Then
				tmp = CType(x, pscCand)
				If Not tmp.Equals(Me) Then
					tmp.Selected = False
				End If
			End If
		Next
	End Sub
	Private Sub cleanHistory()
		mvarQuery = ""
		mvarLastAlive = Nothing
		mvarLastAlarm = ""
		mvarIsAlarm = False
		mvarStatus = mitaStatus.Starting
	End Sub
	Private Sub sendMessage(ByVal message As String)
		mvarIsAlarm = True
		Led.ledColor = problemcolor
		Led.ledStatus = ledModus.ledBlink
		Selected = True
		sendContent()
		RaiseEvent mitaAlarm(Me, mvarDbLine.alarm)
	End Sub
	Private Sub doTheQuery()
		If mvarQuery.Equals("") Then buildQuery()
		Dim idbc As OdbcCommand = odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		idbc.CommandText = mvarQuery
		Try
			odbc_connection.Open()
			reader = idbc.ExecuteReader()
			If reader.Read Then
				mvarDbLine.loginHost = CStr(reader.Item("loginhost"))
				mvarDbLine.loginApp = CStr(reader.Item("loginapp"))
				mvarDbLine.loginId = CStr(reader.Item("loginid"))
				mvarDbLine.sapSystemID = CStr(reader.Item("sapsystemid"))
				If Not IsDBNull(reader.Item("logintime")) Then
					mvarDbLine.loginTime = CDate(reader.Item("logintime"))
				Else
					mvarDbLine.loginTime = Nothing
				End If
				If Not IsDBNull(reader.Item("logouttime")) Then
					mvarDbLine.logoutTime = CDate(reader.Item("logouttime"))
				Else
					mvarDbLine.logoutTime = Nothing
				End If
				If Not IsDBNull(reader.Item("alarm")) Then
					mvarDbLine.alarm = CStr(reader.Item("alarm"))
				Else
					mvarDbLine.alarm = ""
				End If
				If Not IsDBNull(reader.Item("command")) Then
					mvarDbLine.command = CStr(reader.Item("command"))
				Else
					mvarDbLine.command = ""
				End If
				If Not IsDBNull(reader.Item("orderno")) Then
					mvarDbLine.orderNo = CStr(reader.Item("orderno"))
				Else
					mvarDbLine.orderNo = ""
				End If
				If Not IsDBNull(reader.Item("lastorder")) Then
					mvarDbLine.lastOrder = CStr(reader.Item("lastorder"))
				Else
					mvarDbLine.lastOrder = ""
				End If
				If Not IsDBNull(reader.Item("alive")) Then
					mvarDbLine.alive = CDate(reader.Item("alive"))
				Else
					mvarDbLine.alive = Nothing
				End If
				If Not IsDBNull(reader.Item("processid")) Then
					mvarDbLine.processId = CStr(reader.Item("processid"))
				Else
					mvarDbLine.processId = Nothing
				End If
				If Not IsDBNull(reader.Item("caption")) Then
					mvarDbLine.caption = CStr(reader.Item("caption"))
				Else
					mvarDbLine.caption = Nothing
				End If
			End If
			reader.Close()
			idbc.Dispose()
		Catch
			MsgBox(Err.Description & vbCrLf & mvarQuery)
		End Try
		odbc_connection.Close()
	End Sub
	Private Sub buildQuery()
		mvarQuery = "SELECT * FROM " & mvarTable
		mvarQuery = mvarQuery & " WHERE loginhost = '" & mvarHost & "'"
		mvarQuery = mvarQuery & " AND loginid = " & CInt(mvarID)
		mvarQuery = mvarQuery & " And loginapp = '" & mvarProgram & "'"
	End Sub
	Private Function connectionOpenOdbc() As Boolean
		Try
			With odbc_connection
				.ConnectionString = mvarConnectString
				.ConnectionTimeout = 5
				.Open()
				.Close()
			End With
			Return True
		Catch
			Return False
		End Try
	End Function
#End Region
#Region "Local Event Handlers"
	Private Sub lblHost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblHost.Click
		Selected = True
		sendContent()
	End Sub
	Private Sub lblID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblID.Click
		Selected = True
		sendContent()
	End Sub
	Private Sub lblProgram_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblProgram.Click
		Selected = True
		sendContent()
	End Sub
	Private Sub Led_Click1(ByVal sender As Object, ByVal e As System.EventArgs) Handles Led.Click
		Selected = True
		sendContent()
	End Sub
	Private Sub pscCand_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Click
		Selected = True
		sendContent()
	End Sub
	Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
		If Not mvarIsAlarm Then Led.ledStatus = ledModus.ledBlink
		Timer1.Enabled = False
		Timer2.Enabled = True
	End Sub
	Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
		Dim cmpDate As Date = Nothing
		Timer2.Enabled = False
		doTheQuery()
		If mvarStatus = mitaStatus.Starting Then
			mvarStatus = mitaStatus.Waiting
			mvarWaitStartTime = Now
			mvarLastAlarm = mvarDbLine.alarm
			mvarStartAlive = mvarDbLine.alive
			mvarLastAlive = Nothing
			If Selected Then sendContent()
		End If
		If mvarDbLine.loginTime = cmpDate Then
			Led.ledStatus = ledModus.ledOff
			Led.ledColor = offlineColor
			mvarStatus = mitaStatus.NotRunning
			If Selected Then sendContent()
		ElseIf mvarDbLine.loginTime <> cmpDate And mvarDbLine.logoutTime <> cmpDate Then
			Led.ledStatus = ledModus.ledOff
			Led.ledColor = offlineColor
			mvarStatus = mitaStatus.NotRunning
			If Selected Then sendContent()
		Else
			If mvarStatus = mitaStatus.NotRunning Then
				Led.ledColor = testColor
				mvarStatus = mitaStatus.Waiting
				mvarWaitStartTime = Now
				mvarLastAlarm = mvarDbLine.alarm
				mvarStartAlive = mvarDbLine.alive
				mvarLastAlive = Nothing
				If Selected Then sendContent()
			End If
			If Not mvarIsAlarm Then
				Led.ledStatus = ledModus.ledOn
				If mvarLastAlarm <> mvarDbLine.alarm Then
					mvarLastAlarm = mvarDbLine.alarm
					mvarStatus = mitaStatus.Problem
					sendMessage(mvarDbLine.alarm)
				ElseIf Not IsNothing(mvarDbLine.alive) Then
					If mvarDbLine.alive <> mvarStartAlive Then
						If mvarDbLine.alive <> mvarLastAlive Then
							If mvarStatus = mitaStatus.Waiting Then
								mvarStatus = mitaStatus.Running
								Led.ledColor = okColor
							End If
							mvarLastAlive = mvarDbLine.alive
							If Selected Then sendContent()
							mvarAliveTime = Now
						ElseIf Now > DateAdd(DateInterval.Second, TimeOut, mvarAliveTime) Then
							mvarStatus = mitaStatus.TimedOut
							sendMessage("TimeOUT")
						End If
					Else
						If mvarStatus = mitaStatus.Waiting Then
							If Now > DateAdd(DateInterval.Second, TimeOut, mvarWaitStartTime) Then
								mvarStatus = mitaStatus.TimedOut
								sendMessage("TimeOUT")
							End If
						End If
					End If
				End If
			End If
		End If
		Timer1.Enabled = True
	End Sub
#End Region

	Protected Overrides Sub Finalize()
		MyBase.Finalize()
	End Sub
End Class
