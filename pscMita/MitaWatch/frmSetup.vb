Imports pscLedEx.pscLed
Imports System.Reflection
Public Class frmSetup
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
	Friend WithEvents btnDone As System.Windows.Forms.Button
	Friend WithEvents btnAbort As System.Windows.Forms.Button
	Friend WithEvents cbBeep As System.Windows.Forms.ComboBox
	Friend WithEvents Label1 As System.Windows.Forms.Label
	Friend WithEvents Label2 As System.Windows.Forms.Label
	Friend WithEvents cbSound As System.Windows.Forms.ComboBox
	Friend WithEvents Label3 As System.Windows.Forms.Label
	Friend WithEvents txtTimeOut As System.Windows.Forms.TextBox
	Friend WithEvents Label4 As System.Windows.Forms.Label
	Friend WithEvents txtInterval As System.Windows.Forms.TextBox
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmSetup))
		Me.btnDone = New System.Windows.Forms.Button
		Me.btnAbort = New System.Windows.Forms.Button
		Me.cbBeep = New System.Windows.Forms.ComboBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.cbSound = New System.Windows.Forms.ComboBox
		Me.Label3 = New System.Windows.Forms.Label
		Me.txtTimeOut = New System.Windows.Forms.TextBox
		Me.txtInterval = New System.Windows.Forms.TextBox
		Me.Label4 = New System.Windows.Forms.Label
		Me.SuspendLayout()
		'
		'btnDone
		'
		Me.btnDone.DialogResult = System.Windows.Forms.DialogResult.OK
		Me.btnDone.Location = New System.Drawing.Point(152, 156)
		Me.btnDone.Name = "btnDone"
		Me.btnDone.Size = New System.Drawing.Size(76, 20)
		Me.btnDone.TabIndex = 0
		Me.btnDone.Text = "Apply"
		'
		'btnAbort
		'
		Me.btnAbort.DialogResult = System.Windows.Forms.DialogResult.Cancel
		Me.btnAbort.Location = New System.Drawing.Point(70, 156)
		Me.btnAbort.Name = "btnAbort"
		Me.btnAbort.Size = New System.Drawing.Size(76, 20)
		Me.btnAbort.TabIndex = 1
		Me.btnAbort.Text = "Abort"
		'
		'cbBeep
		'
		Me.cbBeep.Location = New System.Drawing.Point(90, 18)
		Me.cbBeep.Name = "cbBeep"
		Me.cbBeep.Size = New System.Drawing.Size(138, 21)
		Me.cbBeep.TabIndex = 2
		'
		'Label1
		'
		Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.Location = New System.Drawing.Point(16, 18)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(56, 18)
		Me.Label1.TabIndex = 3
		Me.Label1.Text = "Audio"
		'
		'Label2
		'
		Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.Location = New System.Drawing.Point(16, 50)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(56, 18)
		Me.Label2.TabIndex = 5
		Me.Label2.Text = "Sound"
		'
		'cbSound
		'
		Me.cbSound.Location = New System.Drawing.Point(90, 50)
		Me.cbSound.Name = "cbSound"
		Me.cbSound.Size = New System.Drawing.Size(138, 21)
		Me.cbSound.TabIndex = 4
		'
		'Label3
		'
		Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(16, 82)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(100, 18)
		Me.Label3.TabIndex = 6
		Me.Label3.Text = "Timeout / sec"
		'
		'txtTimeOut
		'
		Me.txtTimeOut.Location = New System.Drawing.Point(184, 82)
		Me.txtTimeOut.Name = "txtTimeOut"
		Me.txtTimeOut.Size = New System.Drawing.Size(44, 20)
		Me.txtTimeOut.TabIndex = 7
		Me.txtTimeOut.Text = ""
		'
		'txtInterval
		'
		Me.txtInterval.Location = New System.Drawing.Point(184, 113)
		Me.txtInterval.Name = "txtInterval"
		Me.txtInterval.Size = New System.Drawing.Size(44, 20)
		Me.txtInterval.TabIndex = 9
		Me.txtInterval.Text = ""
		'
		'Label4
		'
		Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.Location = New System.Drawing.Point(16, 114)
		Me.Label4.Name = "Label4"
		Me.Label4.Size = New System.Drawing.Size(162, 18)
		Me.Label4.TabIndex = 8
		Me.Label4.Text = "Sample Interval / sec"
		'
		'frmSetup
		'
		Me.AcceptButton = Me.btnDone
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.CancelButton = Me.btnAbort
		Me.ClientSize = New System.Drawing.Size(244, 193)
		Me.Controls.Add(Me.txtInterval)
		Me.Controls.Add(Me.Label4)
		Me.Controls.Add(Me.txtTimeOut)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.cbSound)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.cbBeep)
		Me.Controls.Add(Me.btnAbort)
		Me.Controls.Add(Me.btnDone)
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.Name = "frmSetup"
		Me.Text = "pscMitaWatch Setup"
		Me.ResumeLayout(False)

	End Sub

#End Region
	Dim mvarAudio As ledAudio
	Dim mvarLed As pscLedEx.pscLed
	Dim mvarSound As String
	Dim mvarSampleInterval As Integer
	Dim mvarTimeOut As Integer
	Property Timeout() As Integer
		Get
			Return mvarTimeOut
		End Get
		Set(ByVal Value As Integer)
			mvarTimeOut = Value
			txtTimeOut.Text = CStr(mvarTimeOut)
		End Set
	End Property
	Property SampleInterval() As Integer
		Get
			Return mvarSampleInterval
		End Get
		Set(ByVal Value As Integer)
			mvarSampleInterval = Value
			txtInterval.Text = CStr(mvarSampleInterval)
		End Set
	End Property
	Property Sound() As String
		Get
			Return mvarSound
		End Get
		Set(ByVal Value As String)
			mvarSound = Value
			If cbSound.Items.Count > 0 Then
				cbSound.SelectedItem = mvarSound
			End If
		End Set
	End Property
	Property Audio() As ledAudio
		Get
			Return mvarAudio
		End Get
		Set(ByVal Value As ledAudio)
			mvarAudio = Value
			If cbBeep.Items.Count > 0 Then
				cbBeep.SelectedIndex = mvarAudio
			End If
		End Set
	End Property
	WriteOnly Property Led() As pscLedEx.pscLed
		Set(ByVal Value As pscLedEx.pscLed)
			mvarLed = Value
			fillCombo(cbBeep, New ledAudio)
			getFiles(cbSound, ".wav")
			If Not IsNothing(mvarAudio) Then cbBeep.SelectedItem = mvarAudio
			If Not IsNothing(mvarSound) Then cbSound.SelectedItem = mvarSound
		End Set
	End Property
	Private Sub btnDone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDone.Click
		Me.Hide()
	End Sub

	Private Sub btnAbort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAbort.Click
		Me.Hide()
	End Sub

	Private Sub fillCombo(ByRef cb As Windows.Forms.ComboBox, ByRef test As Object)
		Dim i As Integer
		cb.Items.Clear()
		Dim items() As String = System.Enum.GetNames(test.GetType)
		For i = 0 To UBound(items)
			cb.Items.Add(items(i))
		Next
		cb.SelectedIndex = 0
	End Sub
	Private Sub getFiles(ByVal cb As Windows.Forms.ComboBox, ByVal pattern As String)
		Dim N() As String = [Assembly].GetExecutingAssembly().GetManifestResourceNames
		Dim i As Integer
		cb.Items.Clear()
		For i = 0 To UBound(N)
			If N(i).EndsWith(pattern) Then
				cb.Items.Add(N(i))
			End If
		Next i
		If cb.Items.Count > 0 Then cb.SelectedIndex = 0
	End Sub

	Private Sub frmSetup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		mitaShared.RestPos(Me, False)
	End Sub

	Private Sub frmSetup_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
		mitaShared.SavePos(Me)
	End Sub

	Private Sub cbBeep_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbBeep.SelectedIndexChanged
		mvarAudio = CType(sender, ComboBox).SelectedIndex
	End Sub

	Private Sub cbSound_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbSound.SelectedIndexChanged
		mvarSound = CType(sender, ComboBox).SelectedItem
	End Sub

	Private Sub txtTimeOut_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTimeOut.TextChanged
		mvarTimeOut = Val(CType(sender, TextBox).Text)
	End Sub

	Private Sub txtInterval_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInterval.TextChanged
		mvarSampleInterval = Val(CType(sender, TextBox).Text)
	End Sub
End Class
