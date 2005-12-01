Option Strict On
Option Explicit On 
Module modMsgTimer

	Public message As String
	Public tim As Short
	Public Titel As String

	Delegate Sub TimerProcDelegate(ByVal hWnd As Integer, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTime As Integer)
	Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Integer, ByVal nIDEvent As Integer, ByVal uElapse As Integer, ByVal lpTimerFunc As TimerProcDelegate) As Integer
	Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Integer, ByVal nIDEvent As Integer) As Integer
	Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
	Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Integer) As Integer
	'// Message we receive telling us to close the message box
	Public Const NV_CLOSEMSGBOX As Integer = &H5000
	Public Sub TimerProc(ByVal hWnd As Integer, ByVal uMsg As Integer, ByVal idEvent As Integer, ByVal dwTime As Integer)
		'// this is a callback function.  This means that windows "calls back" to this function
		'// when it's time for the timer event to fire
		'// first thing we do is kill the timer so that no other timer events will fire
		KillTimer(hWnd, idEvent)
		'// select the type of manipulation that we want to perform
		Dim hMessageBox As Integer
		Select Case idEvent
			Case NV_CLOSEMSGBOX		 '// we want to close this messagebox after 4 seconds
				'// find the messagebox window
				'// change the text to whatever the title of the message box is
				hMessageBox = FindWindow("#32770", Titel)
				'// if we found it make sure it has the keyboard focus and then send it an enter to dismiss it
				If hMessageBox > 0 Then
					Call SetForegroundWindow(hMessageBox)
					'// this will result in the default option being chosen
					System.Windows.Forms.SendKeys.Send("{enter}")
				End If
		End Select
	End Sub
End Module