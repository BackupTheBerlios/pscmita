Option Strict On
Option Explicit On 
Public Class CMitaMsg
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
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.components = New System.ComponentModel.Container
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CMitaMsg))
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
		'
		'CMitaMsg
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.ClientSize = New System.Drawing.Size(420, 150)
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.Location = New System.Drawing.Point(155, 284)
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.Name = "CMitaMsg"
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "pscMitaMsg"

	End Sub
#End Region
	Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Integer, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Integer) As Integer
	Private Const MB_SYSTEMMODAL As Integer = &H1000
	Private Const MB_ICONEXCLAMATION As Integer = &H30
	Private mvarWaitTime As Short = 0
	Public Sub popupMessage(ByRef userMessage As String, ByRef userTitle As String)
		message = userMessage
		Titel = userTitle
		Beep()
		If mvarWaitTime = 0 Then
			MsgBox(message, MsgBoxStyle.OKOnly Or MsgBoxStyle.Exclamation, Titel)
		Else
			SetTimer(Me.Handle.ToInt32, NV_CLOSEMSGBOX, 1000 * mvarWaitTime, AddressOf TimerProc)
			MessageBox(Me.Handle.ToInt32, message, Titel, MB_SYSTEMMODAL Or MB_ICONEXCLAMATION)
		End If
	End Sub
	Public Property waitTime() As Short
		Get
			Return mvarWaitTime
		End Get
		Set(ByVal Value As Short)
			mvarWaitTime = Value
		End Set
	End Property
End Class