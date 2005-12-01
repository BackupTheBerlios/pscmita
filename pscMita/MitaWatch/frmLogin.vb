Option Strict Off
Option Explicit On
Friend Class frmLogin
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
	Public WithEvents txtBase As System.Windows.Forms.TextBox
	Public WithEvents txtUserName As System.Windows.Forms.TextBox
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents txtPassword As System.Windows.Forms.TextBox
	Public WithEvents _lblLabels_2 As System.Windows.Forms.Label
	Public WithEvents _lblLabels_0 As System.Windows.Forms.Label
	Public WithEvents _lblLabels_1 As System.Windows.Forms.Label
	Public WithEvents lblLabels As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLogin))
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
		Me.txtBase = New System.Windows.Forms.TextBox
		Me.txtUserName = New System.Windows.Forms.TextBox
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.txtPassword = New System.Windows.Forms.TextBox
		Me._lblLabels_2 = New System.Windows.Forms.Label
		Me._lblLabels_0 = New System.Windows.Forms.Label
		Me._lblLabels_1 = New System.Windows.Forms.Label
		Me.lblLabels = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(Me.components)
		CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'txtBase
		'
		Me.txtBase.AcceptsReturn = True
		Me.txtBase.AutoSize = False
		Me.txtBase.BackColor = System.Drawing.SystemColors.Window
		Me.txtBase.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtBase.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtBase.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtBase.Location = New System.Drawing.Point(87, 12)
		Me.txtBase.MaxLength = 0
		Me.txtBase.Name = "txtBase"
		Me.txtBase.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtBase.Size = New System.Drawing.Size(155, 23)
		Me.txtBase.TabIndex = 6
		Me.txtBase.Text = ""
		'
		'txtUserName
		'
		Me.txtUserName.AcceptsReturn = True
		Me.txtUserName.AutoSize = False
		Me.txtUserName.BackColor = System.Drawing.SystemColors.Window
		Me.txtUserName.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUserName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtUserName.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUserName.Location = New System.Drawing.Point(86, 37)
		Me.txtUserName.MaxLength = 0
		Me.txtUserName.Name = "txtUserName"
		Me.txtUserName.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUserName.Size = New System.Drawing.Size(155, 23)
		Me.txtUserName.TabIndex = 1
		Me.txtUserName.Text = ""
		'
		'cmdOK
		'
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Location = New System.Drawing.Point(33, 96)
		Me.cmdOK.Name = "cmdOK"
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.Size = New System.Drawing.Size(76, 26)
		Me.cmdOK.TabIndex = 4
		Me.cmdOK.Text = "OK"
		'
		'cmdCancel
		'
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Location = New System.Drawing.Point(140, 96)
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.Size = New System.Drawing.Size(76, 26)
		Me.cmdCancel.TabIndex = 5
		Me.cmdCancel.Text = "Cancel"
		'
		'txtPassword
		'
		Me.txtPassword.AcceptsReturn = True
		Me.txtPassword.AutoSize = False
		Me.txtPassword.BackColor = System.Drawing.SystemColors.Window
		Me.txtPassword.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtPassword.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtPassword.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtPassword.Location = New System.Drawing.Point(86, 63)
		Me.txtPassword.MaxLength = 0
		Me.txtPassword.Name = "txtPassword"
		Me.txtPassword.PasswordChar = Microsoft.VisualBasic.ChrW(42)
		Me.txtPassword.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtPassword.Size = New System.Drawing.Size(155, 23)
		Me.txtPassword.TabIndex = 3
		Me.txtPassword.Text = ""
		'
		'_lblLabels_2
		'
		Me._lblLabels_2.BackColor = System.Drawing.SystemColors.Control
		Me._lblLabels_2.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLabels_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblLabels_2.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLabels.SetIndex(Me._lblLabels_2, CType(2, Short))
		Me._lblLabels_2.Location = New System.Drawing.Point(8, 13)
		Me._lblLabels_2.Name = "_lblLabels_2"
		Me._lblLabels_2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLabels_2.Size = New System.Drawing.Size(72, 18)
		Me._lblLabels_2.TabIndex = 7
		Me._lblLabels_2.Text = "&Data Base:"
		'
		'_lblLabels_0
		'
		Me._lblLabels_0.BackColor = System.Drawing.SystemColors.Control
		Me._lblLabels_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLabels_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblLabels_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLabels.SetIndex(Me._lblLabels_0, CType(0, Short))
		Me._lblLabels_0.Location = New System.Drawing.Point(7, 38)
		Me._lblLabels_0.Name = "_lblLabels_0"
		Me._lblLabels_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLabels_0.Size = New System.Drawing.Size(72, 18)
		Me._lblLabels_0.TabIndex = 0
		Me._lblLabels_0.Text = "&User Name:"
		'
		'_lblLabels_1
		'
		Me._lblLabels_1.BackColor = System.Drawing.SystemColors.Control
		Me._lblLabels_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._lblLabels_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._lblLabels_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblLabels.SetIndex(Me._lblLabels_1, CType(1, Short))
		Me._lblLabels_1.Location = New System.Drawing.Point(7, 64)
		Me._lblLabels_1.Name = "_lblLabels_1"
		Me._lblLabels_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._lblLabels_1.Size = New System.Drawing.Size(72, 18)
		Me._lblLabels_1.TabIndex = 2
		Me._lblLabels_1.Text = "&Password:"
		'
		'frmLogin
		'
		Me.AcceptButton = Me.cmdOK
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.CancelButton = Me.cmdCancel
		Me.ClientSize = New System.Drawing.Size(250, 127)
		Me.ControlBox = False
		Me.Controls.Add(Me.txtBase)
		Me.Controls.Add(Me.txtUserName)
		Me.Controls.Add(Me.cmdOK)
		Me.Controls.Add(Me.cmdCancel)
		Me.Controls.Add(Me.txtPassword)
		Me.Controls.Add(Me._lblLabels_2)
		Me.Controls.Add(Me._lblLabels_0)
		Me.Controls.Add(Me._lblLabels_1)
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Location = New System.Drawing.Point(336, 339)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "frmLogin"
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "psc Mita Login"
		CType(Me.lblLabels, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)

	End Sub
#End Region 

	Public loginSucceeded As Boolean
	
	Public connectString As String
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		loginSucceeded = False
		Me.Hide()
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		connectString = "DSN=" & txtBase.Text & ";UID=" & txtUserName.Text & ";PWD=" & txtPassword.Text & ";"
		loginSucceeded = True
		Me.Hide()
	End Sub
	
	Private Sub frmLogin_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		RestPos(Me)
	End Sub
	
		Private Sub frmLogin_Closing(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		Dim Cancel As Short = eventArgs.Cancel
		SavePos(Me)
		eventArgs.Cancel = Cancel
	End Sub

End Class