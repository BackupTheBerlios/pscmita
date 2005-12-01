Imports System.Data.Odbc
Module modMitaXpress
	Public mitaConnect As New pscMitaConnect.CMitaConnect
	Public mitaShared As New pscMitaShared.CMitaShared
	Public mitaData As New pscMitaData.CMitaData
	Public mitaSystem As New pscMitaSapSystem.CMitaSapSystem
	Sub main()
		Dim result As DialogResult
		Dim start As New frmStart
		Dim systems() As String
		Dim indexes() As Integer
		Dim count As Integer
		Dim i As Integer
		Dim hostName(0) As String
		mitaData.registryApplication = "PSC\pscMitaXPress"
		Dim loginCls As New pscMitaLogin.CLogin
		loginCls.registryKey = "PSC\pscMitaLogin"
		loginCls.popUp = True
		If Not loginCls.doLogin Then End
		mitaData.createUser = loginCls.user
		mitaConnect.dbConnectString = loginCls.connectString
		loginCls.Dispose()
		mitaShared.dataSet = mitaData
		mitaShared.connectSet = mitaConnect
		mitaShared.systemSet = mitaSystem
		hostName(0) = System.Net.Dns.GetHostName
		start.hosts = hostName
		start.Text = "pscMitaXPress - Start Program"
		start.cbHost.SelectedIndex = 0
		readDBSap(mitaConnect.odbc_connection, systems, indexes, count)
		start.sapsystems = systems
		start.bldProgs()
		start.optCustom.Checked = True
		result = start.ShowDialog
		If result = DialogResult.OK Then
			If Not RunProcess(start.program, start.arguments) Then
				MsgBox("Start Successless")
			End If
		End If
		start.Dispose()
		End
	End Sub
	Public Function RunProcess(ByVal strProgramName As String, ByVal strArgs As String) As Boolean
		Try
			Dim proc As Process = New Process
			proc.StartInfo.FileName = strProgramName
			If Not IsNothing(strArgs) Then
				proc.StartInfo.Arguments = strArgs
			End If
			Return proc.Start()
		Catch
			Return False
		End Try
	End Function
	'Private Sub login()
	'	Dim frmLoginInst As New pscMitaLogin.CLogin
	'	frmLoginInst.txtUserName.Text = GetSetting("PSC\pscMitaLogin", "lastUser", "login")
	'	frmLoginInst.txtPassword.Text = ""
	'	frmLoginInst.txtBase.Text = GetSetting("PSC\pscMitaLogin", "lastUser", "database")
	'	If frmLoginInst.ShowDialog() = DialogResult.Cancel Then End
	'	mitaConnect.dbConnectString = "DSN=" & frmLoginInst.txtBase.Text & ";UID=" & frmLoginInst.txtUserName.Text & ";PWD=" & frmLoginInst.txtPassword.Text & ";"
	'	frmLoginInst.Dispose()
	'End Sub
	Public Function readDBSap(ByVal odbc_connection As OdbcConnection, ByRef cSap() As String, ByRef cSapIndex() As Integer, ByRef cCount As Integer) As Boolean
		Dim query As String
		Dim a As String
		readDBSap = False
		cCount = -1
		Dim idbc As OdbcCommand = odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		query = "SELECT sapsystemid, sapname FROM pscsapsystems"
		query = query & " WHERE activ = 'Y'"
		query = query & " ORDER BY sapsystemid;"
		idbc.CommandText = query
		Try
			odbc_connection.Open()
			reader = idbc.ExecuteReader()
			While reader.Read
				cCount = cCount + 1
				ReDim Preserve cSap(cCount)
				ReDim Preserve cSapIndex(cCount)
				cSapIndex(cCount) = reader.GetDouble(0)
				cSap(cCount) = reader.GetString(1)
				readDBSap = True
			End While
		Catch
			MsgBox(Err.Description & vbCrLf & query)
		End Try
		reader.Close()
		idbc.Dispose()
		odbc_connection.Close()
	End Function
End Module
