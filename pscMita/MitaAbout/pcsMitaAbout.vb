Imports Microsoft.VisualBasic
Imports System.Drawing
Public Class CMitaAbout
	Inherits System.Windows.Forms.Form

	Private mitaSapSystem As pscMitaSapSystem.CMitaSapSystem
	Private mitaData As pscMitaData.CMitaData
	Private mitaShared As pscMitaShared.CMitaShared

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
	Friend WithEvents cmdOk As System.Windows.Forms.Button
	Friend WithEvents lblVersion As System.Windows.Forms.Label
	Friend WithEvents lblProgram As System.Windows.Forms.Label
	Friend WithEvents pictBackground As System.Windows.Forms.PictureBox
	Friend WithEvents pictIcon As System.Windows.Forms.PictureBox
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents Label4 As System.Windows.Forms.Label
	Friend WithEvents Label5 As System.Windows.Forms.Label
	Friend WithEvents lblSapSystem As System.Windows.Forms.Label
	Friend WithEvents LblSapVersion As System.Windows.Forms.Label
	Friend WithEvents Label6 As System.Windows.Forms.Label
	Friend WithEvents lblHost As System.Windows.Forms.Label
	Friend WithEvents lblHostID As System.Windows.Forms.Label
	Friend WithEvents Label10 As System.Windows.Forms.Label
	Friend WithEvents Label11 As System.Windows.Forms.Label
	Friend WithEvents lblMail As System.Windows.Forms.Label
	Public WithEvents ToolTip1 As System.Windows.Forms.ToolTip
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CMitaAbout))
		Me.pictBackground = New System.Windows.Forms.PictureBox
		Me.cmdOk = New System.Windows.Forms.Button
		Me.lblVersion = New System.Windows.Forms.Label
		Me.lblProgram = New System.Windows.Forms.Label
		Me.pictIcon = New System.Windows.Forms.PictureBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label5 = New System.Windows.Forms.Label
		Me.lblSapSystem = New System.Windows.Forms.Label
		Me.LblSapVersion = New System.Windows.Forms.Label
		Me.Label6 = New System.Windows.Forms.Label
		Me.lblMail = New System.Windows.Forms.Label
		Me.lblHost = New System.Windows.Forms.Label
		Me.lblHostID = New System.Windows.Forms.Label
		Me.Label10 = New System.Windows.Forms.Label
		Me.Label11 = New System.Windows.Forms.Label
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
		Me.SuspendLayout()
		'
		'pictBackground
		'
		Me.pictBackground.Image = CType(resources.GetObject("pictBackground.Image"), System.Drawing.Image)
		Me.pictBackground.Location = New System.Drawing.Point(0, 0)
		Me.pictBackground.Name = "pictBackground"
		Me.pictBackground.Size = New System.Drawing.Size(576, 403)
		Me.pictBackground.TabIndex = 0
		Me.pictBackground.TabStop = False
		'
		'cmdOk
		'
		Me.cmdOk.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.cmdOk.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOk.Location = New System.Drawing.Point(135, 252)
		Me.cmdOk.Name = "cmdOk"
		Me.cmdOk.Size = New System.Drawing.Size(162, 34)
		Me.cmdOk.TabIndex = 1
		Me.cmdOk.Text = "ok"
		'
		'lblVersion
		'
		Me.lblVersion.BackColor = System.Drawing.Color.Silver
		Me.lblVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblVersion.Location = New System.Drawing.Point(206, 108)
		Me.lblVersion.Name = "lblVersion"
		Me.lblVersion.Size = New System.Drawing.Size(62, 16)
		Me.lblVersion.TabIndex = 2
		Me.lblVersion.Text = "Version"
		'
		'lblProgram
		'
		Me.lblProgram.BackColor = System.Drawing.Color.Silver
		Me.lblProgram.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblProgram.ForeColor = System.Drawing.Color.Maroon
		Me.lblProgram.Location = New System.Drawing.Point(114, 76)
		Me.lblProgram.Name = "lblProgram"
		Me.lblProgram.Size = New System.Drawing.Size(154, 26)
		Me.lblProgram.TabIndex = 3
		Me.lblProgram.Text = "Program"
		'
		'pictIcon
		'
		Me.pictIcon.BackColor = System.Drawing.Color.Silver
		Me.pictIcon.Location = New System.Drawing.Point(52, 76)
		Me.pictIcon.Name = "pictIcon"
		Me.pictIcon.Size = New System.Drawing.Size(48, 48)
		Me.pictIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
		Me.pictIcon.TabIndex = 4
		Me.pictIcon.TabStop = False
		'
		'Label1
		'
		Me.Label1.BackColor = System.Drawing.Color.Silver
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.ForeColor = System.Drawing.Color.FromArgb(CType(128, Byte), CType(64, Byte), CType(64, Byte))
		Me.Label1.Location = New System.Drawing.Point(20, 38)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(80, 32)
		Me.Label1.TabIndex = 5
		Me.Label1.Text = "pscMita"
		'
		'Label2
		'
		Me.Label2.BackColor = System.Drawing.Color.Silver
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(104, 38)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(282, 32)
		Me.Label2.TabIndex = 6
		Me.Label2.Text = "Interface to SAP IS / MAM"
		'
		'Label3
		'
		Me.Label3.BackColor = System.Drawing.Color.Silver
		Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(114, 108)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(62, 16)
		Me.Label3.TabIndex = 7
		Me.Label3.Text = "Version:"
		'
		'Label4
		'
		Me.Label4.BackColor = System.Drawing.Color.Silver
		Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Location = New System.Drawing.Point(114, 138)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(98, 16)
		Me.Label4.TabIndex = 8
		Me.Label4.Text = "SAP System:"
		'
		'Label5
		'
		Me.Label5.BackColor = System.Drawing.Color.Silver
		Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.Location = New System.Drawing.Point(268, 138)
		Me.Label5.Name = "Label5"
		Me.Label5.Size = New System.Drawing.Size(98, 16)
		Me.Label5.TabIndex = 9
		Me.Label5.Text = "SAP Version:"
		'
		'lblSapSystem
		'
		Me.lblSapSystem.BackColor = System.Drawing.Color.Silver
		Me.lblSapSystem.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSapSystem.Location = New System.Drawing.Point(206, 138)
		Me.lblSapSystem.Name = "lblSapSystem"
		Me.lblSapSystem.Size = New System.Drawing.Size(62, 16)
		Me.lblSapSystem.TabIndex = 10
		Me.lblSapSystem.Text = "Version"
		'
		'LblSapVersion
		'
		Me.LblSapVersion.BackColor = System.Drawing.Color.Silver
		Me.LblSapVersion.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.LblSapVersion.Location = New System.Drawing.Point(360, 138)
		Me.LblSapVersion.Name = "LblSapVersion"
		Me.LblSapVersion.Size = New System.Drawing.Size(62, 16)
		Me.LblSapVersion.TabIndex = 11
		Me.LblSapVersion.Text = "Version"
		'
		'Label6
		'
		Me.Label6.BackColor = System.Drawing.Color.Silver
		Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.Location = New System.Drawing.Point(114, 186)
		Me.Label6.Name = "Label6"
		Me.Label6.Size = New System.Drawing.Size(234, 20)
		Me.Label6.TabIndex = 12
		Me.Label6.Text = "Copyright (C) Peter Schlang"
		'
		'lblMail
		'
		Me.lblMail.BackColor = System.Drawing.Color.Silver
		Me.lblMail.Cursor = System.Windows.Forms.Cursors.Hand
		Me.lblMail.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblMail.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
		Me.lblMail.Location = New System.Drawing.Point(114, 205)
		Me.lblMail.Name = "lblMail"
		Me.lblMail.Size = New System.Drawing.Size(134, 19)
		Me.lblMail.TabIndex = 13
		Me.lblMail.Text = "psc@my-tools4you.de"
		Me.ToolTip1.SetToolTip(Me.lblMail, "Send Mail")
		'
		'lblHost
		'
		Me.lblHost.BackColor = System.Drawing.Color.Silver
		Me.lblHost.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblHost.Location = New System.Drawing.Point(360, 156)
		Me.lblHost.Name = "lblHost"
		Me.lblHost.Size = New System.Drawing.Size(62, 16)
		Me.lblHost.TabIndex = 17
		Me.lblHost.Text = "Version"
		'
		'lblHostID
		'
		Me.lblHostID.BackColor = System.Drawing.Color.Silver
		Me.lblHostID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblHostID.Location = New System.Drawing.Point(206, 156)
		Me.lblHostID.Name = "lblHostID"
		Me.lblHostID.Size = New System.Drawing.Size(62, 16)
		Me.lblHostID.TabIndex = 16
		Me.lblHostID.Text = "Version"
		'
		'Label10
		'
		Me.Label10.BackColor = System.Drawing.Color.Silver
		Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label10.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.Label10.Location = New System.Drawing.Point(268, 156)
		Me.Label10.Name = "Label10"
		Me.Label10.Size = New System.Drawing.Size(98, 16)
		Me.Label10.TabIndex = 15
		Me.Label10.Text = "Host:"
		'
		'Label11
		'
		Me.Label11.BackColor = System.Drawing.Color.Silver
		Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label11.Location = New System.Drawing.Point(114, 156)
		Me.Label11.Name = "Label11"
		Me.Label11.Size = New System.Drawing.Size(98, 16)
		Me.Label11.TabIndex = 14
		Me.Label11.Text = "Application ID:"
		'
		'CMitaAbout
		'
		Me.AcceptButton = Me.cmdOk
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.CancelButton = Me.cmdOk
		Me.ClientSize = New System.Drawing.Size(432, 303)
		Me.ControlBox = False
		Me.Controls.Add(Me.lblHost)
		Me.Controls.Add(Me.lblHostID)
		Me.Controls.Add(Me.Label10)
		Me.Controls.Add(Me.Label11)
		Me.Controls.Add(Me.lblMail)
		Me.Controls.Add(Me.Label6)
		Me.Controls.Add(Me.LblSapVersion)
		Me.Controls.Add(Me.lblSapSystem)
		Me.Controls.Add(Me.Label5)
		Me.Controls.Add(Me.Label4)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.pictIcon)
		Me.Controls.Add(Me.lblProgram)
		Me.Controls.Add(Me.lblVersion)
		Me.Controls.Add(Me.cmdOk)
		Me.Controls.Add(Me.pictBackground)
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "CMitaAbout"
		Me.Text = "About"
		Me.ResumeLayout(False)

	End Sub

#End Region

	Private Sub pscMitaAbout_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		Dim name As System.Reflection.AssemblyName = System.Reflection.Assembly.GetEntryAssembly.GetName()
		lblVersion.Text = name.Version.Major & "." & name.Version.Minor & "." & name.Version.Revision
		Text = "pscMita Package - About " & name.Name
		Dim tmp As Bitmap = Me.Icon.ToBitmap
		pictIcon.Image = tmp
		lblProgram.Text = name.Name
		mitaShared.appName = "PSC\" & name.Name
		mitaShared.RestPos(Me, False)
	End Sub
	Public WriteOnly Property dataSet()
		Set(ByVal Value)
			mitaData = Value
		End Set
	End Property
	Public WriteOnly Property sharedSet() As pscMitaShared.CMitaShared
		Set(ByVal Value As pscMitaShared.CMitaShared)
			mitaShared = Value
			If mitaData.createID <> -1 Then
				lblHostID.Text = mitaData.createID
			Else
				lblHostID.Text = "N/A"
			End If
			lblHost.Text = mitaData.createHost
		End Set
	End Property
	Public WriteOnly Property systemSet() As pscMitaSapSystem.CMitaSapSystem
		Set(ByVal Value As pscMitaSapSystem.CMitaSapSystem)
			mitaSapSystem = Value
			If IsNothing(mitaSapSystem.sapSystemVERSIONNAME) Then
				LblSapVersion.Text = "N/A"
			Else
				LblSapVersion.Text = mitaSapSystem.sapSystemVERSIONNAME
			End If
			If IsNothing(mitaSapSystem.sapSystemNAME) Then
				lblSapSystem.Text = "N/A"
			Else
				lblSapSystem.Text = mitaSapSystem.sapSystemNAME
			End If
		End Set
	End Property

	'Public Sub setVersion(ByVal version As String)
	'	LblSapVersion.Text = version
	'End Sub
	'Public Sub setSystem(ByVal system As String)
	'	lblSapSystem.Text = system
	'End Sub
	'Public Sub setHostID(ByVal id As String)
	'	lblHostID.Text = id
	'End Sub
	'Public Sub setHost(ByVal host As String)
	'	lblHost.Text = host
	'End Sub

	Private Sub lblMail_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblMail.Click
		Dim bAns As Boolean = True
		Dim sParams As String
		sParams = "mailto:" & lblMail.Text
		sParams = sParams & "?subject=pscMita"
		Try
			System.Diagnostics.Process.Start(sParams)
		Catch
			bAns = False
		End Try
	End Sub

	Private Sub cmdOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOk.Click
		mitaShared.SavePos(Me)
		Me.Dispose()
	End Sub

End Class
