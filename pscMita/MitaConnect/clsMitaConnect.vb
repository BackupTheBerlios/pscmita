Option Strict On
Option Explicit On 
Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.Odbc
Public Class CMitaConnect
	Private mitaData As pscMitaData.CMitaData
	Private mvarDbConnectString As String
	Private mvarOdbc_connection As New OdbcConnection
	Private savStatus As System.Data.ConnectionState

	Private mvarDataBaseType As String = "oracle"
	Private transact As OdbcTransaction
	Private connectStack(10) As Boolean
	Private stackIndex As Integer = -1
	Private mvarErrorDescription As String = ""
	Public Event sqlError(ByVal description As String, ByVal query As String, ByVal title As String)
	Public ReadOnly Property errorDescription() As String
		Get
			Return mvarErrorDescription
		End Get
	End Property
	Public Function stringForDb(ByRef buf1 As String) As String
		Dim buf2 As String
		buf2 = Replace(buf1, "'", " ")
		buf2 = Replace(buf2, "@", " ")
		buf2 = Replace(buf2, "´", " ")
		buf2 = Replace(buf2, "`", " ")
		stringForDb = buf2
	End Function

	Public Function connectionCloseOdbc() As Boolean
		Try
			mvarOdbc_connection.Close()
			Return True
		Catch
			mvarErrorDescription = Err.Description
			Return False
		End Try
	End Function

	Public Function connectionOpenOdbc() As Boolean
		Try
			With mvarOdbc_connection
				.ConnectionString = mvarDbConnectString
				.ConnectionTimeout = 5
				.Open()
			End With
			Return True
		Catch
			mvarErrorDescription = Err.Description
			Return False
		End Try
	End Function

	Public Function connectionTest(ByRef connectString As String) As String
		Dim conn As New OdbcConnection
		On Error GoTo isErr
		With conn
			.ConnectionString = connectString
			.ConnectionTimeout = 5
			.Open()
		End With
		connectionTest = ""
Exx:
		On Error Resume Next
		conn.Close()
		conn.Dispose()
		conn = Nothing
		Exit Function
isErr:
		connectionTest = Err.Description
		mvarErrorDescription = Err.Description
		Resume Exx
	End Function

	Public Function execSQL(ByRef query As String) As Boolean
		Dim idbc As OdbcCommand = mvarOdbc_connection.CreateCommand()
		savStatus = mvarOdbc_connection.State
		execSQL = False
		idbc.CommandText = query
		Try
			If savStatus = ConnectionState.Closed Then mvarOdbc_connection.Open()
			transact = mvarOdbc_connection.BeginTransaction(System.Data.IsolationLevel.ReadCommitted)
			idbc.Transaction = transact
			idbc.ExecuteNonQuery()
			transact.Commit()
			idbc.Dispose()
			execSQL = True
		Catch
			mvarErrorDescription = Err.Description
			transact.Rollback()
			RaiseEvent sqlError(mvarErrorDescription, query, "execSQL")
		End Try
		If savStatus = ConnectionState.Closed Then mvarOdbc_connection.Close()
	End Function

	Public Function queryExist(ByRef query As String, ByRef result As Boolean) As Boolean
		Dim idbc As OdbcCommand = mvarOdbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		savStatus = mvarOdbc_connection.State
		queryExist = False
		idbc.CommandText = query
		Try
			If savStatus = ConnectionState.Closed Then mvarOdbc_connection.Open()
			reader = idbc.ExecuteReader()
			result = reader.HasRows
			reader.Close()
			idbc.Dispose()
			queryExist = True
		Catch
			mvarErrorDescription = Err.Description
			RaiseEvent sqlError(mvarErrorDescription, query, "queryExist")
		End Try
		If savStatus = ConnectionState.Closed Then mvarOdbc_connection.Close()
	End Function

	Public Function queryNumber(ByRef query As String, ByRef number As Integer) As Boolean
		Dim idbc As OdbcCommand = mvarOdbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		savStatus = mvarOdbc_connection.State
		idbc.CommandText = query
		queryNumber = False
		Try
			If savStatus = ConnectionState.Closed Then mvarOdbc_connection.Open()
			reader = idbc.ExecuteReader()
			If reader.Read Then
				queryNumber = True
				number = CInt(reader.GetValue(0))
			End If
			reader.Close()
			idbc.Dispose()
		Catch
			mvarErrorDescription = Err.Description
			RaiseEvent sqlError(mvarErrorDescription, query, "queryNumber")
		End Try
		If savStatus = ConnectionState.Closed Then mvarOdbc_connection.Close()
	End Function
	Public Function queryString(ByRef query As String, ByRef target As String) As Boolean
		Dim idbc As OdbcCommand = mvarOdbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		savStatus = mvarOdbc_connection.State
		queryString = False
		idbc.CommandText = query
		Try
			If savStatus = ConnectionState.Closed Then mvarOdbc_connection.Open()
			reader = idbc.ExecuteReader()
			queryString = False
			If reader.Read Then
				queryString = True
				target = reader.GetString(0)
			End If
			reader.Close()
			idbc.Dispose()
		Catch
			mvarErrorDescription = Err.Description
			RaiseEvent sqlError(mvarErrorDescription, query, "queryString")
		End Try
		If savStatus = ConnectionState.Closed Then mvarOdbc_connection.Close()
	End Function

	Public Function tableExists(ByVal tableName As String) As Boolean
		Dim query As String = "SELECT * FROM all_objects"
		Dim idbc As OdbcCommand = mvarOdbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		savStatus = mvarOdbc_connection.State
		query = query & " WHERE object_type IN ('TABLE','VIEW')"
		query = query & " AND object_name = '" & UCase$(tableName) & "';"
		tableExists = False
		idbc.CommandText = query
		Try
			If savStatus = ConnectionState.Closed Then mvarOdbc_connection.Open()
			reader = idbc.ExecuteReader()
			If reader.Read Then
				Dim x As String = reader.Item(0).ToString
				tableExists = True
			End If
			reader.Close()
			idbc.Dispose()
		Catch
			mvarErrorDescription = Err.Description
			RaiseEvent sqlError(mvarErrorDescription, query, "tableExists")
		End Try
		If savStatus = ConnectionState.Closed Then mvarOdbc_connection.Close()
	End Function
	Public Function dbDate(ByVal dat As Date) As String
		Return "to_date('" & Format(dat, "yyyyMMdd") & "', 'YYYYMMDD')"
	End Function

	Public Sub connectPush()
		stackIndex = stackIndex + 1
		connectStack(stackIndex) = (mvarOdbc_connection.State = ConnectionState.Open)
		If connectStack(stackIndex) Then mvarOdbc_connection.Close()
	End Sub

	Public Sub connectPop()
		Try
			If connectStack(stackIndex) Then
				mvarOdbc_connection.Open()
			Else
				mvarOdbc_connection.Close()
			End If
		Catch
			mvarErrorDescription = Err.Description
			RaiseEvent sqlError(mvarErrorDescription, "", "connectPop")
		End Try
		stackIndex = stackIndex - 1
	End Sub
	Property dataSet() As pscMitaData.CMitaData
		Get
			Return mitaData
		End Get
		Set(ByVal Value As pscMitaData.CMitaData)
			If IsNothing(mitaData) Then mitaData = New pscMitaData.CMitaData
			mitaData = Value
			'mvarDataBaseType = mitaData.sapSystemDATABASETYPE
		End Set
	End Property
	Property dbConnectString() As String
		Get
			Return mvarDbConnectString
		End Get
		Set(ByVal Value As String)
			mvarDbConnectString = Value
			mvarOdbc_connection.ConnectionString = mvarDbConnectString
		End Set
	End Property
	Property odbc_connection() As OdbcConnection
		Get
			Return mvarOdbc_connection
		End Get
		Set(ByVal Value As OdbcConnection)
			mvarOdbc_connection = Value
		End Set
	End Property

End Class