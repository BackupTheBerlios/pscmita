Option Strict Off
Option Explicit On 
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Data.Odbc
Public Class CLogin
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
	Public WithEvents txtBase As System.Windows.Forms.TextBox
	Public WithEvents txtUserName As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents txtPassword As System.Windows.Forms.TextBox
	Public WithEvents _lblLabels_2 As System.Windows.Forms.Label
	Public WithEvents _lblLabels_0 As System.Windows.Forms.Label
	Public WithEvents _lblLabels_1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CLogin))
		Me.txtBase = New System.Windows.Forms.TextBox
		Me.txtUserName = New System.Windows.Forms.TextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.txtPassword = New System.Windows.Forms.TextBox
		Me._lblLabels_2 = New System.Windows.Forms.Label
		Me._lblLabels_0 = New System.Windows.Forms.Label
		Me._lblLabels_1 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		'
		'txtBase
		'
		Me.txtBase.AcceptsReturn = True
		Me.txtBase.AutoSize = False
		Me.txtBase.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtBase.Location = New System.Drawing.Point(92, 10)
		Me.txtBase.MaxLength = 0
		Me.txtBase.Name = "txtBase"
		Me.txtBase.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtBase.TabIndex = 6
		Me.txtBase.Text = ""
		'
		'txtUserName
		'
		Me.txtUserName.AcceptsReturn = True
		Me.txtUserName.AutoSize = False
		Me.txtUserName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUserName.Location = New System.Drawing.Point(92, 43)
		Me.txtUserName.MaxLength = 0
		Me.txtUserName.Name = "txtUserName"
		Me.txtUserName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUserName.TabIndex = 1
		Me.txtUserName.Text = ""
		'
		'cmdOK
		'
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.DialogResult = System.Windows.Forms.DialogResult.OK
		Me.cmdOK.Location = New System.Drawing.Point(204, 75)
		Me.cmdOK.Name = "cmdOK"
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabIndex = 4
		Me.cmdOK.Text = "OK"
		'
		'cmdCancel
		'
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.cmdCancel.Location = New System.Drawing.Point(204, 42)
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabIndex = 5
		Me.cmdCancel.Text = "Cancel"
		'
		'txtPassword
		'
		Me.txtPassword.AcceptsReturn = True
		Me.txtPassword.AutoSize = False
		Me.txtPassword.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPassword.Location = New System.Drawing.Point(92, 76)
		Me.txtPassword.MaxLength = 0
		Me.txtPassword.Name = "txtPassword"
		Me.txtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
		Me.txtPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPassword.TabIndex = 3
		Me.txtPassword.Text = ""
		'
		'_lblLabels_2
		'
		Me._lblLabels_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLabels_2.Location = New System.Drawing.Point(4, 10)
		Me._lblLabels_2.Name = "_lblLabels_2"
		Me._lblLabels_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLabels_2.Size = New System.Drawing.Size(100, 20)
		Me._lblLabels_2.TabIndex = 7
		Me._lblLabels_2.Text = "&Data Base:"
		Me._lblLabels_2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		'
		'_lblLabels_0
		'
		Me._lblLabels_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLabels_0.Location = New System.Drawing.Point(4, 43)
		Me._lblLabels_0.Name = "_lblLabels_0"
		Me._lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLabels_0.Size = New System.Drawing.Size(100, 20)
		Me._lblLabels_0.TabIndex = 0
		Me._lblLabels_0.Text = "&User Name:"
		Me._lblLabels_0.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		'
		'_lblLabels_1
		'
		Me._lblLabels_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLabels_1.Location = New System.Drawing.Point(4, 76)
		Me._lblLabels_1.Name = "_lblLabels_1"
		Me._lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLabels_1.Size = New System.Drawing.Size(100, 20)
		Me._lblLabels_1.TabIndex = 2
		Me._lblLabels_1.Text = "&Password:"
		Me._lblLabels_1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
		'
		'CLogin
		'
		Me.AcceptButton = Me.cmdOK
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.CancelButton = Me.cmdCancel
		Me.ClientSize = New System.Drawing.Size(292, 109)
		Me.ControlBox = False
		Me.Controls.Add(Me.txtBase)
		Me.Controls.Add(Me.txtUserName)
		Me.Controls.Add(Me.txtPassword)
		Me.Controls.Add(Me.cmdOK)
		Me.Controls.Add(Me.cmdCancel)
		Me.Controls.Add(Me._lblLabels_2)
		Me.Controls.Add(Me._lblLabels_0)
		Me.Controls.Add(Me._lblLabels_1)
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "CLogin"
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "psc Mita Login"
		Me.ResumeLayout(False)

	End Sub
#End Region
	Private mvarRegistryKey As String
	Private mvarConnectString As String = ""
	Private mitaShared As New pscMitaShared.CMitaShared
	Private mvarUser As String
	Private mvarDataBase As String
	Private mvarPwd As String
	Private mvarCmd As String
	Private mvarPopUp As Boolean = False
	Public WriteOnly Property popUp() As Boolean
		Set(ByVal Value As Boolean)
			mvarPopUp = Value
		End Set
	End Property
	Public ReadOnly Property dataBase() As String
		Get
			Return mvarDataBase
		End Get
	End Property
	Public ReadOnly Property connectString() As String
		Get
			Return mvarConnectString
		End Get
	End Property
	Public ReadOnly Property user() As String
		Get
			Return mvarUser
		End Get
	End Property
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		mvarConnectString = ""
		Me.Hide()
	End Sub

	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		mvarConnectString = "DSN=" & txtBase.Text & ";UID=" & txtUserName.Text & ";PWD=" & txtPassword.Text & ";"
		Dim a As String = connectionTest(mvarConnectString)
		If a <> "" Then
			MsgBox(a)
			Exit Sub
		End If
		Me.Hide()
	End Sub

	Private Sub frmLogin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		centerMe()
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
	Public WriteOnly Property registryKey() As String
		Set(ByVal Value As String)
			mvarRegistryKey = Value
		End Set
	End Property
	Public WriteOnly Property commandLine() As String
		Set(ByVal Value As String)
			mvarCmd = Value
		End Set
	End Property
	Public Function doLogin() As Boolean
		Dim x() As String
		Dim y() As String
		Dim i As Integer
		Dim a As String

		mvarUser = GetSetting(mvarRegistryKey, "lastUser", "login")
		mvarDataBase = GetSetting(mvarRegistryKey, "lastUser", "database")
		If Not mvarPopUp Then mvarPwd = mitaShared.DoXor(GetSetting(mvarRegistryKey, "lastUser", "pwd"))
		x = Split(mvarCmd, "-")
		For i = 0 To UBound(x)
			If x(i) <> "" Then
				y = Split(x(i), " ")
				Select Case UCase(y(0))
					Case "UID", "U"
						mvarUser = y(1)
					Case "DSN", "D"
						mvarDataBase = y(1)
					Case "PWD", "P"
						mvarPwd = mitaShared.DoXor(y(1))
				End Select
			End If
		Next i
		mvarConnectString = "DSN=" & mvarDataBase & ";UID=" & mvarUser & ";PWD=" & mvarPwd & ";"
		If mvarUser = "" Or mvarPwd = "" Or mvarDataBase = "" Then
			txtUserName.Text = mvarUser
			txtPassword.Text = mvarPwd
			txtBase.Text = mvarDataBase
			If ShowDialog() = DialogResult.Cancel Then Return False
			a = connectionTest(mvarConnectString)
			If a <> "" Then
				MsgBox(a, MsgBoxStyle.Critical Or MsgBoxStyle.OKOnly)
				mvarConnectString = ""
				Return False
			End If
			mvarUser = txtUserName.Text
			mvarDataBase = txtBase.Text
			mvarPwd = txtPassword.Text
		End If
		SaveSetting(mvarRegistryKey, "lastUser", "login", mvarUser)
		SaveSetting(mvarRegistryKey, "lastUser", "database", mvarDataBase)
		SaveSetting(mvarRegistryKey, "lastUser", "pwd", mitaShared.DoXor(mvarPwd))
		Return True
	End Function
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
End Class