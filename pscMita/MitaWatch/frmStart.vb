Public Class frmStart
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
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents cbHost As System.Windows.Forms.ComboBox
	Friend WithEvents cbProgram As System.Windows.Forms.ComboBox
	Friend WithEvents btnDone As System.Windows.Forms.Button
	Friend WithEvents cmdStart As System.Windows.Forms.Button
	Friend WithEvents txtList As System.Windows.Forms.TextBox
	Friend WithEvents cbSAP As System.Windows.Forms.ComboBox
	Friend WithEvents Label4 As System.Windows.Forms.Label
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
	Friend WithEvents optProd As System.Windows.Forms.RadioButton
	Friend WithEvents optTest As System.Windows.Forms.RadioButton
	Friend WithEvents optDevelop As System.Windows.Forms.RadioButton
	Friend WithEvents chkTrace As System.Windows.Forms.CheckBox
	Friend WithEvents chkProfile As System.Windows.Forms.CheckBox
	Friend WithEvents chkMini As System.Windows.Forms.CheckBox
	Friend WithEvents chkGarbage As System.Windows.Forms.CheckBox
	Friend WithEvents chkSame As System.Windows.Forms.CheckBox
	Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
	Friend WithEvents chkAds As System.Windows.Forms.CheckBox
	Friend WithEvents chkInfo As System.Windows.Forms.CheckBox
	Friend WithEvents chkOrder As System.Windows.Forms.CheckBox
	Friend WithEvents chkErrors As System.Windows.Forms.CheckBox
	Friend WithEvents Label5 As System.Windows.Forms.Label
	Friend WithEvents Label6 As System.Windows.Forms.Label
	Friend WithEvents optRegistry As System.Windows.Forms.RadioButton
	Friend WithEvents optCustom As System.Windows.Forms.RadioButton
	Friend WithEvents txtPool As System.Windows.Forms.TextBox
	Friend WithEvents grpOptions As System.Windows.Forms.GroupBox
	Friend WithEvents chkSave As System.Windows.Forms.CheckBox
	Friend WithEvents optNone As System.Windows.Forms.RadioButton
	Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
	Friend WithEvents optRestore As System.Windows.Forms.RadioButton
	Friend WithEvents optCenter As System.Windows.Forms.RadioButton
	Friend WithEvents optHide As System.Windows.Forms.RadioButton
	Friend WithEvents chkTaskbar As System.Windows.Forms.CheckBox
	Friend WithEvents chkTray As System.Windows.Forms.CheckBox
	Friend WithEvents Label7 As System.Windows.Forms.Label
	Friend WithEvents txtAlive As System.Windows.Forms.TextBox
	Friend WithEvents chkBatch As System.Windows.Forms.CheckBox
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.cbHost = New System.Windows.Forms.ComboBox
		Me.cbProgram = New System.Windows.Forms.ComboBox
		Me.btnDone = New System.Windows.Forms.Button
		Me.cmdStart = New System.Windows.Forms.Button
		Me.grpOptions = New System.Windows.Forms.GroupBox
		Me.chkBatch = New System.Windows.Forms.CheckBox
		Me.txtAlive = New System.Windows.Forms.TextBox
		Me.Label7 = New System.Windows.Forms.Label
		Me.txtPool = New System.Windows.Forms.TextBox
		Me.Label5 = New System.Windows.Forms.Label
		Me.GroupBox3 = New System.Windows.Forms.GroupBox
		Me.chkErrors = New System.Windows.Forms.CheckBox
		Me.chkOrder = New System.Windows.Forms.CheckBox
		Me.chkInfo = New System.Windows.Forms.CheckBox
		Me.chkAds = New System.Windows.Forms.CheckBox
		Me.txtList = New System.Windows.Forms.TextBox
		Me.cbSAP = New System.Windows.Forms.ComboBox
		Me.Label4 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.GroupBox2 = New System.Windows.Forms.GroupBox
		Me.optNone = New System.Windows.Forms.RadioButton
		Me.optProd = New System.Windows.Forms.RadioButton
		Me.optTest = New System.Windows.Forms.RadioButton
		Me.optDevelop = New System.Windows.Forms.RadioButton
		Me.chkTrace = New System.Windows.Forms.CheckBox
		Me.chkProfile = New System.Windows.Forms.CheckBox
		Me.chkMini = New System.Windows.Forms.CheckBox
		Me.chkGarbage = New System.Windows.Forms.CheckBox
		Me.chkSame = New System.Windows.Forms.CheckBox
		Me.Label6 = New System.Windows.Forms.Label
		Me.optRegistry = New System.Windows.Forms.RadioButton
		Me.optCustom = New System.Windows.Forms.RadioButton
		Me.chkSave = New System.Windows.Forms.CheckBox
		Me.GroupBox1 = New System.Windows.Forms.GroupBox
		Me.chkTray = New System.Windows.Forms.CheckBox
		Me.chkTaskbar = New System.Windows.Forms.CheckBox
		Me.optHide = New System.Windows.Forms.RadioButton
		Me.optCenter = New System.Windows.Forms.RadioButton
		Me.optRestore = New System.Windows.Forms.RadioButton
		Me.grpOptions.SuspendLayout()
		Me.GroupBox3.SuspendLayout()
		Me.GroupBox2.SuspendLayout()
		Me.GroupBox1.SuspendLayout()
		Me.SuspendLayout()
		'
		'Label1
		'
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(12, 16)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(56, 20)
		Me.Label1.TabIndex = 0
		Me.Label1.Text = "Host"
		'
		'Label2
		'
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(12, 48)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(62, 20)
		Me.Label2.TabIndex = 1
		Me.Label2.Text = "Program"
		'
		'cbHost
		'
		Me.cbHost.Location = New System.Drawing.Point(94, 16)
		Me.cbHost.Name = "cbHost"
		Me.cbHost.Size = New System.Drawing.Size(130, 21)
		Me.cbHost.TabIndex = 2
		'
		'cbProgram
		'
		Me.cbProgram.Location = New System.Drawing.Point(94, 48)
		Me.cbProgram.Name = "cbProgram"
		Me.cbProgram.Size = New System.Drawing.Size(130, 21)
		Me.cbProgram.TabIndex = 3
		'
		'btnDone
		'
		Me.btnDone.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.btnDone.Location = New System.Drawing.Point(240, 48)
		Me.btnDone.Name = "btnDone"
		Me.btnDone.Size = New System.Drawing.Size(104, 20)
		Me.btnDone.TabIndex = 4
		Me.btnDone.Text = "Abort"
		'
		'cmdStart
		'
		Me.cmdStart.DialogResult = System.Windows.Forms.DialogResult.OK
		Me.cmdStart.Enabled = False
		Me.cmdStart.Location = New System.Drawing.Point(240, 16)
		Me.cmdStart.Name = "cmdStart"
		Me.cmdStart.Size = New System.Drawing.Size(104, 20)
		Me.cmdStart.TabIndex = 5
		Me.cmdStart.Text = "Start"
		'
		'grpOptions
		'
		Me.grpOptions.Controls.Add(Me.chkBatch)
		Me.grpOptions.Controls.Add(Me.txtAlive)
		Me.grpOptions.Controls.Add(Me.Label7)
		Me.grpOptions.Controls.Add(Me.txtPool)
		Me.grpOptions.Controls.Add(Me.Label5)
		Me.grpOptions.Controls.Add(Me.GroupBox3)
		Me.grpOptions.Controls.Add(Me.txtList)
		Me.grpOptions.Controls.Add(Me.cbSAP)
		Me.grpOptions.Controls.Add(Me.Label4)
		Me.grpOptions.Controls.Add(Me.Label3)
		Me.grpOptions.Controls.Add(Me.GroupBox2)
		Me.grpOptions.Controls.Add(Me.chkTrace)
		Me.grpOptions.Controls.Add(Me.chkProfile)
		Me.grpOptions.Controls.Add(Me.chkMini)
		Me.grpOptions.Controls.Add(Me.chkGarbage)
		Me.grpOptions.Controls.Add(Me.chkSame)
		Me.grpOptions.Enabled = False
		Me.grpOptions.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.grpOptions.Location = New System.Drawing.Point(14, 128)
		Me.grpOptions.Name = "grpOptions"
		Me.grpOptions.Size = New System.Drawing.Size(332, 262)
		Me.grpOptions.TabIndex = 6
		Me.grpOptions.TabStop = False
		Me.grpOptions.Text = "Custom Start Options"
		'
		'chkBatch
		'
		Me.chkBatch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkBatch.Location = New System.Drawing.Point(14, 128)
		Me.chkBatch.Name = "chkBatch"
		Me.chkBatch.Size = New System.Drawing.Size(154, 16)
		Me.chkBatch.TabIndex = 31
		Me.chkBatch.Text = "Batch SAP Errors"
		'
		'txtAlive
		'
		Me.txtAlive.Location = New System.Drawing.Point(106, 237)
		Me.txtAlive.Name = "txtAlive"
		Me.txtAlive.Size = New System.Drawing.Size(68, 21)
		Me.txtAlive.TabIndex = 30
		Me.txtAlive.Text = "10"
		'
		'Label7
		'
		Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label7.Location = New System.Drawing.Point(14, 237)
		Me.Label7.Name = "Label7"
		Me.Label7.Size = New System.Drawing.Size(80, 20)
		Me.Label7.TabIndex = 29
		Me.Label7.Text = "Alive"
		'
		'txtPool
		'
		Me.txtPool.Location = New System.Drawing.Point(106, 212)
		Me.txtPool.Name = "txtPool"
		Me.txtPool.Size = New System.Drawing.Size(68, 21)
		Me.txtPool.TabIndex = 28
		Me.txtPool.Text = "psc§sappool"
		'
		'Label5
		'
		Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label5.Location = New System.Drawing.Point(14, 212)
		Me.Label5.Name = "Label5"
		Me.Label5.Size = New System.Drawing.Size(80, 20)
		Me.Label5.TabIndex = 27
		Me.Label5.Text = "Pool"
		'
		'GroupBox3
		'
		Me.GroupBox3.Controls.Add(Me.chkErrors)
		Me.GroupBox3.Controls.Add(Me.chkOrder)
		Me.GroupBox3.Controls.Add(Me.chkInfo)
		Me.GroupBox3.Controls.Add(Me.chkAds)
		Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GroupBox3.Location = New System.Drawing.Point(186, 22)
		Me.GroupBox3.Name = "GroupBox3"
		Me.GroupBox3.Size = New System.Drawing.Size(136, 124)
		Me.GroupBox3.TabIndex = 26
		Me.GroupBox3.TabStop = False
		Me.GroupBox3.Text = "Logging"
		'
		'chkErrors
		'
		Me.chkErrors.Checked = True
		Me.chkErrors.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkErrors.Location = New System.Drawing.Point(12, 92)
		Me.chkErrors.Name = "chkErrors"
		Me.chkErrors.TabIndex = 3
		Me.chkErrors.Tag = "E"
		Me.chkErrors.Text = "Errors"
		'
		'chkOrder
		'
		Me.chkOrder.Location = New System.Drawing.Point(12, 68)
		Me.chkOrder.Name = "chkOrder"
		Me.chkOrder.TabIndex = 2
		Me.chkOrder.Tag = "O"
		Me.chkOrder.Text = "Order"
		'
		'chkInfo
		'
		Me.chkInfo.Location = New System.Drawing.Point(12, 44)
		Me.chkInfo.Name = "chkInfo"
		Me.chkInfo.TabIndex = 1
		Me.chkInfo.Tag = "I"
		Me.chkInfo.Text = "Information"
		'
		'chkAds
		'
		Me.chkAds.Location = New System.Drawing.Point(12, 20)
		Me.chkAds.Name = "chkAds"
		Me.chkAds.TabIndex = 0
		Me.chkAds.Tag = "A"
		Me.chkAds.Text = "Ads"
		'
		'txtList
		'
		Me.txtList.Location = New System.Drawing.Point(106, 187)
		Me.txtList.Name = "txtList"
		Me.txtList.Size = New System.Drawing.Size(68, 21)
		Me.txtList.TabIndex = 25
		Me.txtList.Text = "250"
		'
		'cbSAP
		'
		Me.cbSAP.Location = New System.Drawing.Point(106, 160)
		Me.cbSAP.Name = "cbSAP"
		Me.cbSAP.Size = New System.Drawing.Size(68, 23)
		Me.cbSAP.TabIndex = 24
		'
		'Label4
		'
		Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Location = New System.Drawing.Point(14, 187)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(80, 20)
		Me.Label4.TabIndex = 23
		Me.Label4.Text = "List Entries"
		'
		'Label3
		'
		Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(14, 162)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(90, 20)
		Me.Label3.TabIndex = 22
		Me.Label3.Text = "SAP System"
		'
		'GroupBox2
		'
		Me.GroupBox2.Controls.Add(Me.optNone)
		Me.GroupBox2.Controls.Add(Me.optProd)
		Me.GroupBox2.Controls.Add(Me.optTest)
		Me.GroupBox2.Controls.Add(Me.optDevelop)
		Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GroupBox2.Location = New System.Drawing.Point(186, 154)
		Me.GroupBox2.Name = "GroupBox2"
		Me.GroupBox2.Size = New System.Drawing.Size(136, 100)
		Me.GroupBox2.TabIndex = 21
		Me.GroupBox2.TabStop = False
		Me.GroupBox2.Text = "Modus"
		'
		'optNone
		'
		Me.optNone.Checked = True
		Me.optNone.Location = New System.Drawing.Point(9, 80)
		Me.optNone.Name = "optNone"
		Me.optNone.Size = New System.Drawing.Size(118, 18)
		Me.optNone.TabIndex = 3
		Me.optNone.TabStop = True
		Me.optNone.Text = "Standard"
		'
		'optProd
		'
		Me.optProd.Location = New System.Drawing.Point(9, 60)
		Me.optProd.Name = "optProd"
		Me.optProd.Size = New System.Drawing.Size(118, 18)
		Me.optProd.TabIndex = 2
		Me.optProd.Text = "PRODUCTION"
		'
		'optTest
		'
		Me.optTest.Location = New System.Drawing.Point(9, 40)
		Me.optTest.Name = "optTest"
		Me.optTest.Size = New System.Drawing.Size(106, 18)
		Me.optTest.TabIndex = 1
		Me.optTest.Text = "TEST"
		'
		'optDevelop
		'
		Me.optDevelop.Location = New System.Drawing.Point(9, 20)
		Me.optDevelop.Name = "optDevelop"
		Me.optDevelop.Size = New System.Drawing.Size(106, 18)
		Me.optDevelop.TabIndex = 0
		Me.optDevelop.Text = "DEVELOP"
		'
		'chkTrace
		'
		Me.chkTrace.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkTrace.Location = New System.Drawing.Point(14, 108)
		Me.chkTrace.Name = "chkTrace"
		Me.chkTrace.Size = New System.Drawing.Size(154, 16)
		Me.chkTrace.TabIndex = 20
		Me.chkTrace.Text = "SAP Trace"
		'
		'chkProfile
		'
		Me.chkProfile.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkProfile.Location = New System.Drawing.Point(14, 88)
		Me.chkProfile.Name = "chkProfile"
		Me.chkProfile.Size = New System.Drawing.Size(154, 16)
		Me.chkProfile.TabIndex = 19
		Me.chkProfile.Text = "Profiler"
		'
		'chkMini
		'
		Me.chkMini.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkMini.Location = New System.Drawing.Point(14, 68)
		Me.chkMini.Name = "chkMini"
		Me.chkMini.Size = New System.Drawing.Size(154, 16)
		Me.chkMini.TabIndex = 18
		Me.chkMini.Text = "MiniSAP"
		'
		'chkGarbage
		'
		Me.chkGarbage.Checked = True
		Me.chkGarbage.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkGarbage.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkGarbage.Location = New System.Drawing.Point(14, 48)
		Me.chkGarbage.Name = "chkGarbage"
		Me.chkGarbage.Size = New System.Drawing.Size(154, 16)
		Me.chkGarbage.TabIndex = 17
		Me.chkGarbage.Text = "Remove Garbage"
		'
		'chkSame
		'
		Me.chkSame.Checked = True
		Me.chkSame.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkSame.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSame.Location = New System.Drawing.Point(14, 28)
		Me.chkSame.Name = "chkSame"
		Me.chkSame.Size = New System.Drawing.Size(154, 16)
		Me.chkSame.TabIndex = 16
		Me.chkSame.Text = "Allow Same Version"
		'
		'Label6
		'
		Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label6.Location = New System.Drawing.Point(12, 82)
		Me.Label6.Name = "Label6"
		Me.Label6.Size = New System.Drawing.Size(100, 20)
		Me.Label6.TabIndex = 7
		Me.Label6.Text = "Start Options"
		'
		'optRegistry
		'
		Me.optRegistry.Checked = True
		Me.optRegistry.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optRegistry.Location = New System.Drawing.Point(118, 82)
		Me.optRegistry.Name = "optRegistry"
		Me.optRegistry.Size = New System.Drawing.Size(114, 18)
		Me.optRegistry.TabIndex = 8
		Me.optRegistry.TabStop = True
		Me.optRegistry.Text = "From Registry"
		'
		'optCustom
		'
		Me.optCustom.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.optCustom.Location = New System.Drawing.Point(240, 82)
		Me.optCustom.Name = "optCustom"
		Me.optCustom.Size = New System.Drawing.Size(104, 18)
		Me.optCustom.TabIndex = 9
		Me.optCustom.Text = "Individually"
		'
		'chkSave
		'
		Me.chkSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.chkSave.Location = New System.Drawing.Point(240, 108)
		Me.chkSave.Name = "chkSave"
		Me.chkSave.Size = New System.Drawing.Size(114, 16)
		Me.chkSave.TabIndex = 10
		Me.chkSave.Text = "Save in Registry"
		'
		'GroupBox1
		'
		Me.GroupBox1.Controls.Add(Me.chkTray)
		Me.GroupBox1.Controls.Add(Me.chkTaskbar)
		Me.GroupBox1.Controls.Add(Me.optHide)
		Me.GroupBox1.Controls.Add(Me.optCenter)
		Me.GroupBox1.Controls.Add(Me.optRestore)
		Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.GroupBox1.Location = New System.Drawing.Point(12, 410)
		Me.GroupBox1.Name = "GroupBox1"
		Me.GroupBox1.Size = New System.Drawing.Size(332, 94)
		Me.GroupBox1.TabIndex = 11
		Me.GroupBox1.TabStop = False
		Me.GroupBox1.Text = "Display Options"
		'
		'chkTray
		'
		Me.chkTray.Location = New System.Drawing.Point(186, 46)
		Me.chkTray.Name = "chkTray"
		Me.chkTray.Size = New System.Drawing.Size(138, 18)
		Me.chkTray.TabIndex = 4
		Me.chkTray.Text = "Show in Icon Tray"
		'
		'chkTaskbar
		'
		Me.chkTaskbar.Checked = True
		Me.chkTaskbar.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkTaskbar.Location = New System.Drawing.Point(186, 24)
		Me.chkTaskbar.Name = "chkTaskbar"
		Me.chkTaskbar.Size = New System.Drawing.Size(138, 18)
		Me.chkTaskbar.TabIndex = 3
		Me.chkTaskbar.Text = "Show in Taskbar"
		'
		'optHide
		'
		Me.optHide.Location = New System.Drawing.Point(20, 68)
		Me.optHide.Name = "optHide"
		Me.optHide.Size = New System.Drawing.Size(138, 18)
		Me.optHide.TabIndex = 2
		Me.optHide.Text = "Hide from Screen"
		'
		'optCenter
		'
		Me.optCenter.Location = New System.Drawing.Point(20, 46)
		Me.optCenter.Name = "optCenter"
		Me.optCenter.Size = New System.Drawing.Size(138, 18)
		Me.optCenter.TabIndex = 1
		Me.optCenter.Text = "Center Screen"
		'
		'optRestore
		'
		Me.optRestore.Checked = True
		Me.optRestore.Location = New System.Drawing.Point(20, 24)
		Me.optRestore.Name = "optRestore"
		Me.optRestore.Size = New System.Drawing.Size(138, 18)
		Me.optRestore.TabIndex = 0
		Me.optRestore.TabStop = True
		Me.optRestore.Text = "Restore Location"
		'
		'frmStart
		'
		Me.AcceptButton = Me.cmdStart
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.CancelButton = Me.btnDone
		Me.ClientSize = New System.Drawing.Size(358, 517)
		Me.Controls.Add(Me.GroupBox1)
		Me.Controls.Add(Me.chkSave)
		Me.Controls.Add(Me.optCustom)
		Me.Controls.Add(Me.optRegistry)
		Me.Controls.Add(Me.Label6)
		Me.Controls.Add(Me.grpOptions)
		Me.Controls.Add(Me.cmdStart)
		Me.Controls.Add(Me.btnDone)
		Me.Controls.Add(Me.cbProgram)
		Me.Controls.Add(Me.cbHost)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.Label1)
		Me.Name = "frmStart"
		Me.Text = "pscMitaWatch - Start Program"
		Me.grpOptions.ResumeLayout(False)
		Me.GroupBox3.ResumeLayout(False)
		Me.GroupBox2.ResumeLayout(False)
		Me.GroupBox1.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub

#End Region

	Public Sub bldProgs()
		cbProgram.Items.Clear()
		cbProgram.Items.Add("pscMitaOLP")
		cbProgram.Items.Add("pscMitaINP")
		cbProgram.Items.Add("pscMitaSRV")
		cbProgram.Items.Add("pscMitaERR")
	End Sub
	WriteOnly Property hosts() As String()
		Set(ByVal Value As String())
			Dim i As Integer
			cbHost.Items.Clear()
			If IsNothing(Value) Then Exit Property
			For i = 0 To UBound(Value)
				cbHost.Items.Add(Value(i))
			Next
		End Set
	End Property
	WriteOnly Property sapsystems() As String()
		Set(ByVal Value As String())
			Dim i As Integer
			cbSAP.Items.Clear()
			For i = 0 To UBound(Value)
				cbSAP.Items.Add(Value(i))
			Next
		End Set
	End Property
	Property host() As String
		Get
			Return cbHost.Text
		End Get
		Set(ByVal Value As String)
			cbHost.SelectedItem = Value
		End Set
	End Property
	Property sap() As String
		Get
			Return cbSAP.SelectedItem
		End Get
		Set(ByVal Value As String)
			cbSAP.SelectedItem = Value
		End Set
	End Property
	Property program() As String
		Get
			Return cbProgram.SelectedItem
		End Get
		Set(ByVal Value As String)
			cbProgram.SelectedItem = Value
		End Set
	End Property
	Property arguments() As String
		Get
			If optRegistry.Checked Then Return "" Else Return buildArguments()
		End Get
		Set(ByVal Value As String)
			decodeArguments(Value)
		End Set
	End Property
	Private Sub cmdStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStart.Click, btnDone.Click
		Me.Hide()
	End Sub

	Private Sub frmStart_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		mitaShared.RestPos(Me)
	End Sub

	Private Sub decodeArguments(ByVal cmdLine As String)
		Dim y() As String
		Dim i As Integer
		Dim args() As String = cmdLine.Split(" ")
		For i = 0 To UBound(args)
			y = args(i).Split("=")
			Select Case UCase(y(0))
				Case "/SYS"
					cbSAP.SelectedItem = y(1)
				Case "/LIST"
					txtList.Text = y(1)
				Case "/LOG"
					If InStr(1, y(1), "E", CompareMethod.Text) > 0 Then chkErrors.Checked = True
					If InStr(1, y(1), "A", CompareMethod.Text) > 0 Then chkAds.Checked = True
					If InStr(1, y(1), "O", CompareMethod.Text) > 0 Then chkOrder.Checked = True
					If InStr(1, y(1), "I", CompareMethod.Text) > 0 Then chkInfo.Checked = True
				Case "/SAMEVERSION"
					chkSame.Checked = getBool(y)
				Case "/GARBAGE"
					chkGarbage.Checked = getBool(y)
				Case "/PROD"
					optProd.Checked = True
				Case "/TEST"
					optTest.Checked = True
				Case "/DEVELOP"
					optDevelop.Checked = True
				Case "/MINISAP"
					chkMini.Checked = getBool(y)
				Case "/PROFILE"
					chkProfile.Checked = getBool(y)
				Case "/TRACE"
					chkTrace.Checked = (y(1) = 1)
				Case "/POOL"
					txtPool.Text = y(1)
				Case "/RESTORE"
					optRestore.Checked = True
				Case "/CENTER"
					optCenter.Checked = True
				Case "/HIDE"
					optHide.Checked = True
				Case "/TASKBAR"
					chkTaskbar.Checked = getBool(y)
				Case "/ICONTRAY"
					chkTray.Checked = getBool(y)
				Case "/ALIVE"
					txtAlive.Text = y(1)
				Case "/ERRBATCH"
					chkBatch.Checked = getBool(y)
			End Select
		Next i
	End Sub
	Private Function buildArguments() As String
		Dim res As String = "/SYS="
		res = res & cbSAP.Text
		mitaShared.readDBSapSystemFromName(cbSAP.Text)
		res = res & " /LOG="
		res = res & getTag(chkAds)
		res = res & getTag(chkOrder)
		res = res & getTag(chkInfo)
		res = res & getTag(chkErrors)
		res = res & " /SAMEVERSION=" & getOnOff(chkSame)
		res = res & " /GARBAGE=" & getOnOff(chkGarbage)
		res = res & " /MINISAP=" & getOnOff(chkMini)
		res = res & " /PROFILE=" & getOnOff(chkProfile)
		res = res & " /POOL=" & Replace(txtPool.Text, "§", mitaSystem.sapSystemId)
		res = res & " /ALIVE=" & txtAlive.Text
		res = res & " /ERRBATCH=" & getOnOff(chkBatch)
		res = res & " /LIST=" & txtList.Text
		If optDevelop.Checked Then res = res & " /DEVELOP"
		If optTest.Checked Then res = res & " /TEST"
		If optProd.Checked Then res = res & " /PROD"
		If optRestore.Checked Then res = res & " /RESTORE"
		If optCenter.Checked Then res = res & " /CENTER"
		If optHide.Checked Then res = res & " /HIDE"
		If chkSave.Checked Then res = res & " /SAVE"
		res = res & " /TASKBAR=" & getOnOff(chkTaskbar)
		res = res & " /ICONTRAY=" & getOnOff(chkTray)
		If chkTrace.Checked Then
			res = res & " /TRACE=1"
		Else
			res = res & " /TRACE=0"
		End If
		Return res
	End Function
	Private Function getOnOff(ByVal chk As CheckBox)
		Return IIf(chk.Checked, "On", "Off")
	End Function
	Private Function getTag(ByVal chk As CheckBox) As String
		If chk.Checked = False Then Return "" Else Return chk.Tag.ToString
	End Function

	Private Sub optRegistry_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optRegistry.CheckedChanged
		If optRegistry.Checked Then grpOptions.Enabled = False
	End Sub

	Private Sub optCustom_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles optCustom.CheckedChanged
		If optCustom.Checked Then grpOptions.Enabled = True
	End Sub

	Private Sub frmStart_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		mitaShared.SavePos(Me)
	End Sub
	Private Function getBool(ByVal inp() As String) As Boolean
		If UBound(inp) = 0 Then Return True
		Return Not (StrComp(inp(1), "Off", CompareMethod.Text) = 0)
	End Function

	Private Sub cbHost_Changed(ByVal sender As Object, ByVal e As System.EventArgs) Handles cbHost.SelectedIndexChanged, cbProgram.SelectedIndexChanged
		cmdStart.Enabled = (cbHost.Text <> "") And (cbProgram.Text <> "")
	End Sub

	Private Sub frmStart_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
		If cbSAP.SelectedIndex = -1 Then cbSAP.SelectedIndex = 0
	End Sub

End Class
