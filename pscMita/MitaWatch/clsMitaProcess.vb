Option Strict On
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
Imports System.Runtime.Remoting.RemotingConfiguration
Imports pscMitaRunner
Public Class clsMitaProcess
	Private Enum AgentStatus
		notNunning
		netRunning
		localRunning
	End Enum
	Private mvarHostName As String = Nothing
	Private mvarAgentRunning As AgentStatus = AgentStatus.notNunning
	Private mvarCaption As String = Nothing
	Private mvarProcess As Process
	Private mvarChannel As TcpChannel = New TcpChannel(7501)
	Private mvarClient As System.Type
	Private mvarConnection As String = Nothing
	Private runner As pscMitaRunner.CpscMitaRunner
	Private localRunner As pscMitaRunner.CpscMitaRunner = New pscMitaRunner.CpscMitaRunner
	Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
	Public Overloads Sub Dispose()
		MyBase.Finalize()
	End Sub
	Public WriteOnly Property caption() As String
		Set(ByVal Value As String)
			mvarCaption = Value
		End Set
	End Property
	Public WriteOnly Property hostName() As String
		Set(ByVal Value As String)
			mvarHostName = Value
		End Set
	End Property
	Private Sub startAgent()
		If mvarHostName = System.Net.Dns.GetHostName().ToString Then
			mvarHostName = "localhost"
			'runner = localRunner
			'mvarAgentRunning = AgentStatus.localRunning
			'Else
		Else
			Dim i As Integer = 0
		End If
		Try
			mvarConnection = "tcp://" & mvarHostName & ":" & CStr(7500) & "/pscMitaRunner"
			If mvarAgentRunning <> AgentStatus.netRunning Then
				ChannelServices.RegisterChannel(mvarChannel)
				mvarClient = GetType(pscMitaRunner.CpscMitaRunner)
				RemotingConfiguration.RegisterWellKnownClientType(mvarClient, mvarConnection)
				RemotingConfiguration.CustomErrorsEnabled(False)
				mvarAgentRunning = AgentStatus.netRunning
			End If
			runner = DirectCast(Activator.GetObject(mvarClient, mvarConnection), pscMitaRunner.CpscMitaRunner)
		Catch
			MessageBox.Show("Error!!! " + Err.Description)
		End Try
		'End If
	End Sub
	Public Function startMitaApplication(ByVal app As String, ByVal args As String) As Boolean
		startAgent()
		Try
			If runner.RunProcess(app, args) Then
				Return True
			Else
				Return False
			End If
		Catch
			Return False
		End Try
	End Function
	Public Function shutdownMitaApplication() As Boolean
		Dim cnt As Integer = 0
		startAgent()
		findProcess()
		If runner.ShutDownProcess(mvarProcess) Then
			Do
				Sleep(100)
				If Not runner.RetreiveProcess(mvarProcess) Then Return True
				cnt = cnt + 1
				If cnt = 10 Then Return False
			Loop
		Else
			Return False
		End If
	End Function
	Public Function getProcess() As Process
		startAgent()
		findProcess()
		Return mvarProcess
	End Function
	Public Function killMitaApplication() As Boolean
		startAgent()
		findProcess()
		If runner.KillProcess(mvarProcess) Then
			Return True
		Else
			Return False
		End If
	End Function
	Private Sub findProcess()
		If mvarAgentRunning <> AgentStatus.notNunning And Not IsNothing(mvarCaption) Then
			mvarProcess = runner.findProcess(mvarCaption)
		End If
	End Sub
End Class
