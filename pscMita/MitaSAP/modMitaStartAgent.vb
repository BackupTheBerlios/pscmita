Imports System.Diagnostics
Imports System.Runtime.Serialization.Formatters
Module pscMitaStartAgent
	Public Sub startAgent()
		Dim proc As System.Diagnostics.Process
		Dim processes() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
		For Each proc In processes
			If proc.ProcessName = "pscMitaAgent" Then Exit Sub
		Next
		Dim bAns As Boolean = True
		Try
			System.Diagnostics.Process.Start("pscMitaAgent.exe")
		Catch
			bAns = False
		End Try
	End Sub

End Module
