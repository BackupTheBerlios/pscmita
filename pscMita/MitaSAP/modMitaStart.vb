Imports pscMitaDef.CMitaDef
Module modMitaStart
	Public Structure startupStructure
		Dim myID As String
#If SAP Then
		Dim orderDll As PSCMitaOrder.CMitaOrder
		Dim errorDll As PscMitaError.CMitaError
		Dim workModus As workModi
#End If
		Dim cmd As String
		Dim user As String
		Dim pwd As String
		Dim dataBase As String
		Dim id As Integer
		Dim sapSystemName As String
		Dim sapSystemID As Integer
		Dim sapVersionName As String
		Dim runType As String
		Dim alive As Integer
		Dim garbageRemove As Boolean
		Dim sameVersion As Boolean
		Dim sqlClass As mitaSqlClass
		Dim orderDirectory As String
		Dim miniSAP As Boolean
		Dim maxList As Integer
		Dim hostName As String
		Dim typChar As String
		Dim formCaption As String
		Dim trace As Char
		Dim pool As String
		Dim iconIndex As Integer
		Dim caption As String
		Dim processID As Integer
		Dim profile As Boolean
		Dim taskBar As Boolean
		Dim iconTray As Boolean
		Dim centerMe As Boolean
		Dim hideMe As Boolean
		Dim batchError As Boolean
	End Structure

	Public startupInfo As startupStructure
	Public myInstance As Integer
	Public aliveConnection As New pscMitaConnect.CMitaConnect
	Public aliveQuery As String
	Public aliveCount As Integer

	Public Sub doStartup(ByVal frm As System.Windows.Forms.Form, ByVal sapOrder As PSCMitaOrder.CMitaOrder)
		startupInfo.processID = System.Diagnostics.Process.GetCurrentProcess.Id
		startupInfo.hostName = System.Net.Dns.GetHostName
		mitaData.processID = startupInfo.processID
		frm.Text = mitaShared.getExeName() & startupInfo.formCaption
		myInstance = startupInfo.id + ShowCaption(frm)
		myInstance = mitaShared.generateID(myInstance)
		mitaData.caption = frm.Text
		mitaSystem.runType = startupInfo.typChar & startupInfo.runType
		sapOrder.eventRaise("", mitaEventCodes.programStart)
		startupInfo.sapSystemID = mitaSystem.sapSystemID
	End Sub
	Public Function ShowCaption(ByRef applicationForm As System.Windows.Forms.Form) As Short
		Dim C As String
		Dim t As String
		Dim w As Integer
		Dim z As Short
		Dim p As Short
		C = applicationForm.Text
		t = C & " #1"
		p = Len(t)
		t = t & "  -  Version " & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMajorPart & "." & System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).FileMinorPart
		Dim tmp As New System.Windows.Forms.Form
		tmp.Text = t
		mitaShared.buildVersionInfo(tmp)
		t = tmp.Text
		tmp.Dispose()
		z = 1
		Do
			Mid(t, p, 1) = Trim(Str(z))
			w = FindWindow(vbNullString, t)
			If w = 0 Then Exit Do
			z = z + 1
		Loop
		applicationForm.Text = t
		ShowCaption = z
	End Function
	Public Function tryAlive(ByVal conn As pscMitaConnect.CMitaConnect) As Boolean
		aliveCount = aliveCount + 1
		If aliveCount >= startupInfo.alive Then
			Dim cnn As OdbcConnection = conn.odbc_connection
			Dim idbc As OdbcCommand = cnn.CreateCommand()
			cnn.Open()
			idbc.CommandText = aliveQuery
			idbc.ExecuteNonQuery()
			idbc.Dispose()
			cnn.Close()
			aliveCount = 0
			Return True
		End If
		Return False
	End Function
End Module
