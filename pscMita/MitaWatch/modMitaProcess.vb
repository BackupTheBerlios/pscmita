Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data
Imports System.Runtime.Remoting
Imports System.Diagnostics
Imports System.Threading
Imports System.Runtime.Remoting.Channels.Tcp
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Serialization.Formatters
Imports pscMitaRunner
Public Structure mitaProcess
	Dim hostID As Integer
	Dim hostName As Integer
	Dim agentRunning As Boolean
	Dim capture As String
	Dim processID As Long
End Structure
Module modMitaProcess
	Private runner As New pscMitaRunner.CpscMitaRunner

	Public Sub startAgent(ByVal host As String, ByVal deltaPort As Integer)
		Try
			RemotingConfiguration.RegisterWellKnownClientType(GetType(pscMitaRunner.CpscMitaRunner), _
			"tcp://" & host & ":" & CStr(7500 + deltaPort) & "/pscMitaRunner")
			Dim ch As System.Runtime.Remoting.Channels.Tcp.TcpChannel = _
			New System.Runtime.Remoting.Channels.Tcp.TcpChannel(7501 + deltaPort)
			ChannelServices.RegisterChannel(ch)
			RemotingConfiguration.CustomErrorsEnabled(False)
			runner = New pscMitaRunner.CpscMitaRunner
		Catch
			MessageBox.Show("Error!!! " + Err.Description)
		End Try
	End Sub
	Public Function startMitaApplication(ByVal host As String, ByVal app As String) As Boolean
		Try
			If runner.RunProcess(app, "") Then
				Return True
			Else
				Return False
			End If
		Catch
			Return False
		End Try
	End Function
	Public Function shutdownMitaApplication(ByVal proc As Process) As Boolean
		If runner.ShutDownProcess(proc) Then
			Return True
		Else
			Return False
		End If
	End Function
	Public Function killMitaApplication(ByVal proc As Process) As Boolean
		If runner.KillProcess(proc) Then
			Return True
		Else
			Return False
		End If
	End Function
End Module
