Option Strict On
Option Explicit On 
Imports Microsoft.VisualBasic
Imports System.data.odbc
Imports System.data
Imports pscMitaDef.CMitaDef
Module MitaReadDB
	Friend mitaShared As pscMitaShared.CMitaShared
	Friend mitaData As pscMitaData.CMitaData
	Friend mitaConnect As pscMitaConnect.CMitaConnect
	Friend mitaSystem As pscMitaSapSystem.CMitaSapSystem
	Friend mitaMessage As pscMitaMsg.CMitaMsg
	Public Function readDBCustCombis(ByRef cCombis() As COMBISTRUCT, ByRef cCount As Integer) As Boolean
		Dim query As String
		Dim sb As New System.Text.StringBuilder(1000)
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		readDBCustCombis = False
		cCount = -1
		sb.Append("SELECT combiname, items from " & mitaSystem.tableCustCombis)
		sb.Append(" WHERE activ = 'Y'")
		sb.Append(" ORDER BY combiname;")
		query = sb.ToString()
		idbc.CommandText = query
		On Error GoTo isErr
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			cCount = cCount + 1
			ReDim Preserve cCombis(cCount)
			cCombis(cCount).cName = CStr(reader.Item("combiname"))
			cCombis(cCount).cEntry = CStr(reader.Item("items"))
		End While
		readDBCustCombis = True
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
		Resume Exx
	End Function

	Public Function readDBSapVersions(ByRef versions() As Integer, ByRef cCount As Integer) As Boolean
		Dim query As String
		Dim sb As New System.Text.StringBuilder(1000)
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		readDBSapVersions = False
		cCount = -1
		sb.Append("SELECT DISTINCT sapversionid from " & mitaSystem.tableSapSystems)
		sb.Append(" WHERE activ = 'Y'")
		query = sb.ToString()
		idbc.CommandText = query
		On Error GoTo isErr
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			cCount = cCount + 1
			ReDim Preserve versions(cCount)
			versions(cCount) = CInt(reader.Item("sapversionid"))
		End While
		readDBSapVersions = True
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
#If Not MitaOrder And Not MitaError And Not DLL Then
		MsgBox(mitaData.errorDescription)
		End
#End If
		Resume Exx
	End Function

	Public Function readDBSapSystems(ByRef systems() As Integer, ByRef cCount As Integer) As Boolean
		Dim query As String
		Dim sb As New System.Text.StringBuilder(1000)
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		readDBSapSystems = False
		cCount = -1
		sb.Append("SELECT t1.sapsystemid from " & mitaSystem.tableSapSystems & " t1")
		sb.Append(", " & mitaSystem.tableSapVersions & " t2")
		sb.Append(" WHERE t1.activ = 'Y'")
		sb.Append(" AND t1.sapversionid = t2.sapversionid")
		query = sb.ToString()
		idbc.CommandText = query
		On Error GoTo isErr
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			cCount = cCount + 1
			ReDim Preserve systems(cCount)
			systems(cCount) = CInt(reader.Item("sapsystemid"))
		End While
		readDBSapSystems = True
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
		Resume Exx
	End Function
	Public Function readDBCustQueries(ByRef targetQueries() As String, ByRef targetNames() As String) As Boolean
		Dim query As String
		Dim cnt As Integer
		Dim i As Integer
		Dim sav As String
		Dim found As Boolean
		readDBCustQueries = False
		On Error GoTo isErr
		ReDim targetQueries(0)
		ReDim targetNames(0)
		Dim sb As New System.Text.StringBuilder(1000)
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader

		sb.Append("SELECT pscname, text FROM " & mitaSystem.tableCustQuery)
		sb.Append(" WHERE activ = 'Y'")
		sb.Append(" AND sapsystemid = " & mitaSystem.sapSystemID)
		sb.Append(" AND rfctype = " & mitaSystem.rfcType)
		sb.Append(" AND activ = 'Y'")
		sb.Append(" ORDER BY pscname;")
		query = sb.ToString()
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			cnt = cnt + 1
			ReDim Preserve targetQueries(cnt)
			Dim A$ = CStr(reader.Item("text"))
			targetQueries(cnt) = Replace(A$, "~", "'")
			ReDim Preserve targetNames(cnt)
			targetNames(cnt) = CStr(reader.Item("pscname"))
		End While
		readDBCustQueries = True
		'DB sort is different from VB sort!
		' I arrange in Vb sort order
		Do
			found = False
			For i = 0 To cnt - 1
				If targetNames(i) > targetNames(i + 1) Then
					sav = targetNames(i)
					targetNames(i) = targetNames(i + 1)
					targetNames(i + 1) = sav
					sav = targetQueries(i)
					targetQueries(i) = targetQueries(i + 1)
					targetQueries(i + 1) = sav
					found = True
					Exit For
				End If
			Next i
			If Not found Then Exit Do
		Loop
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
#If Not MitaOrder And Not MitaError And Not DLL Then
		MsgBox(mitaData.errorDescription)
		End
#End If
		Resume Exx
	End Function

	Public Function readDbSqlEntry(ByRef source As String, ByRef targetSQL As String) As Boolean
		Dim queryComment As String
		Dim queryResult As Integer
		Dim queryVersion As Integer
		readDbSqlEntry = readDBCustQuery(source, targetSQL, queryVersion, queryComment)
	End Function


	Public Function readDBCustQuery(ByRef queryName As String, ByRef queryStatement As String, ByRef queryVersion As Integer, ByRef queryComment As String) As Boolean
		Dim query As String
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		Dim sb As New System.Text.StringBuilder(1000)
		readDBCustQuery = False
		sb.Append("SELECT text, version, comments from " & mitaSystem.tableCustQuery)
		sb.Append(" WHERE pscname = '" & queryName & "'")
		sb.Append(" AND sapsystemid = " & mitaSystem.sapSystemID)
		sb.Append(" AND rfctype = " & mitaSystem.rfcType)
		sb.Append(" AND activ = 'Y';")
		query = sb.ToString()
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		queryStatement = ""
		queryComment = ""
		queryVersion = -1
		While reader.Read
			queryStatement = Replace(CStr(reader.Item("text")), "~", "'")
			queryVersion = CInt(reader.Item("version"))
			If IsDBNull(reader.Item("comments")) Then
				queryComment = ""
			Else
				queryComment = CStr(reader.Item("comments"))
			End If
			readDBCustQuery = True
		End While
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
#If Not MitaOrder And Not MitaError And Not DLL Then
		MsgBox(mitaData.errorDescription)
		End
#End If
		Resume Exx
	End Function


	Public Function readDBCustTables(ByRef cTables() As TABLESTRUCT, ByRef cCount As Integer) As Boolean
		Dim query As String
		Dim sb As New System.Text.StringBuilder(1000)
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		readDBCustTables = False
		sb.Append("SELECT DISTINCT tablename FROM " & mitaSystem.tableCustTables)
		sb.Append(" WHERE activ = 'Y'")
		sb.Append(" AND sapsystemid = " & mitaSystem.sapSystemID)
		sb.Append(" AND rfctype = " & mitaSystem.rfcType)
		sb.Append(" ORDER BY tablename;")
		query = sb.ToString()
		cCount = -1
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			cCount = cCount + 1
			ReDim Preserve cTables(cCount)
			cTables(cCount).tName = CStr(reader.Item("tablename"))
			cTables(cCount).tCount = -1
			readDBCustTables = True
		End While
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
#If Not MitaOrder And Not MitaError And Not DLL Then
		MsgBox(mitaData.errorDescription)
		End
#End If
		Resume Exx
	End Function


	Public Function readDBCustTableEntry(ByRef name As String, ByRef field As String, ByRef s2a As Boolean, ByRef targetEntry As String) As Boolean
		Dim query As String
		Dim a As String
		Dim sb As New System.Text.StringBuilder(1000)
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		readDBCustTableEntry = False
		sb.Append("SELECT tleft, tright FROM " & mitaSystem.tableCustTables)
		sb.Append(" WHERE sapsystemid = " & mitaSystem.sapSystemID)
		sb.Append(" AND rfctype = " & mitaSystem.rfcType)
		sb.Append(" AND activ = 'Y'")
		sb.Append(" AND tablename = '" & name & "'")
		If s2a Then
			sb.Append(" AND tleft = '" & field & "'")
		Else
			sb.Append(" AND tright = '" & field & "'")
		End If
		sb.Append(";")
		query = sb.ToString()
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			If s2a Then
				If Not IsDBNull(reader.Item("tright")) Then
					targetEntry = CStr(reader.Item("tright"))
				Else
					targetEntry = ""
				End If
			Else
				targetEntry = CStr(reader.Item("tleft"))
			End If
		End While
		readDBCustTableEntry = True
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
#If Not MitaOrder And Not MitaError And Not DLL Then
		MsgBox(mitaData.errorDescription)
		End
#End If
		Resume Exx
	End Function

	Public Function readDBCustTableEntries(ByRef name As String) As custTableArray
		Dim tleftleft() As String
		Dim tleftright() As String
		Dim trightleft() As String
		Dim trightright() As String
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		Dim query As String
		Dim sb As New System.Text.StringBuilder(1000)
		Dim targetll() As String = Nothing
		Dim targetlr() As String = Nothing
		Dim targetrl() As String = Nothing
		Dim targetrr() As String = Nothing
		Dim cnt As Integer = -1
		Dim A$
		sb.Append("SELECT tleft, tright FROM " & mitaSystem.tableCustTables)
		sb.Append(" WHERE sapsystemid = " & mitaSystem.sapSystemID)
		sb.Append(" AND tablename = '" & name & "'")
		sb.Append(" AND rfctype = " & mitaSystem.rfcType)
		sb.Append(" AND activ = 'Y'")
		query = sb.ToString()
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			cnt = cnt + 1
			ReDim Preserve targetll(cnt)
			ReDim Preserve targetlr(cnt)
			ReDim Preserve targetrl(cnt)
			ReDim Preserve targetrr(cnt)
			If Not IsDBNull(reader.Item("tleft")) Then
				A$ = CStr(reader.Item("tleft"))
			Else
				A$ = ""
			End If
			targetll(cnt) = A$
			targetrl(cnt) = A$
			If Not IsDBNull(reader.Item("tright")) Then
				A$ = CStr(reader.Item("tright"))
			Else
				A$ = ""
			End If
			targetlr(cnt) = A$
			targetrr(cnt) = A$
		End While
		Array.Sort(targetll, targetlr)
		Array.Sort(targetrr, targetrl)
		reader.Close()
		idbc.Dispose()
Exx:
		mitaConnect.odbc_connection.Close()
		Dim newArr As custTableArray
		newArr.tableLeftLeft = targetll
		newArr.tableLeftRight = targetlr
		newArr.tableRightLeft = targetrl
		newArr.tableRightRight = targetrr
		Return newArr
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
		Resume Exx
	End Function


	Public Function readDBCustFields(ByRef cFields() As FIELDSTRUCT, ByRef cCount As Integer, Optional ByVal tmpRfc As Integer = -1) As Boolean
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		Dim query As String
		Dim actRfc As rfcClass
		If tmpRfc = -1 Then
			actRfc = CType(mitaSystem.rfcType, rfcClass)
		Else
			actRfc = CType(tmpRfc, rfcClass)
		End If
		readDBCustFields = False
		cCount = -1
		ReDim cFields(0)
		Dim sb As New System.Text.StringBuilder(1000)

		sb.Append("SELECT field, sapstruct, sapfield, first, length, ftype, recno from " & mitaSystem.tableCustFields)
		sb.Append(" WHERE sapsystemid = " & mitaSystem.sapSystemID)
		sb.Append(" AND rfctype = " & actRfc)
		sb.Append(" AND activ = 'Y'")
		sb.Append(" ORDER BY recno;")
		query = sb.ToString()
		On Error GoTo isErr
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			cCount = cCount + 1
			ReDim Preserve cFields(cCount)
			cFields(cCount).fName = CStr(reader.Item("field"))
			cFields(cCount).fFirst = CInt(reader.Item("first"))
			cFields(cCount).fLength = CInt(reader.Item("length"))
			cFields(cCount).fStructure = CStr(reader.Item("sapstruct"))
			cFields(cCount).fStructureField = CStr(reader.Item("sapfield"))
			cFields(cCount).fIndex = CInt(reader.Item("recno"))
			cFields(cCount).fType = CStr(reader.Item("ftype"))
		End While
		readDBCustFields = True
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
#If Not MitaOrder And Not MitaError And Not DLL Then
		MsgBox(mitaData.errorDescription)
		End
#End If
		Resume Exx
	End Function

	Public Function readDBStructures(ByRef cStruct() As STRUCTSTRUCT, ByRef cCount As Integer, Optional ByVal tmpRfc As Integer = -1) As Boolean
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		Dim query As String
		Dim index As Integer
		Dim actRfc As rfcClass
		If tmpRfc = -1 Then
			actRfc = CType(mitaSystem.rfcType, rfcClass)
		Else
			actRfc = CType(tmpRfc, rfcClass)
		End If
		On Error GoTo isErr
		Dim sb As New System.Text.StringBuilder(1000)
		sb.Append("SELECT sapstruct, recno, length, slevel FROM " & mitaSystem.tableStructures)
		sb.Append(" WHERE sapversionid = " & mitaSystem.sapSystemVERSIONID)
		sb.Append(" AND rfctype = " & actRfc)
		sb.Append(" AND activ = 'Y'")
		sb.Append(" ORDER BY recno DESC;")
		query = sb.ToString()
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		cCount = -1
		If reader.Read Then
			cCount = CInt(reader.Item("recno")) - 1
		End If
		If Not IsNothing(reader) Then reader.Close()
		If cCount > 0 Then
			ReDim cStruct(cCount)
			sb.Remove(0, sb.Length)
			sb.Append("SELECT sapstruct, recno, length, slevel, datatype, rfcfunction, datastruct FROM " & mitaSystem.tableStructures)
			sb.Append(" WHERE sapversionid = " & mitaSystem.sapSystemVERSIONID)
			sb.Append(" AND rfctype = " & actRfc)
			sb.Append(" AND activ = 'Y'")
			sb.Append(" ORDER BY recno")
			query = sb.ToString()
			idbc.CommandText = query
			reader = idbc.ExecuteReader()
			While reader.Read
				index = CInt(reader.Item("recno")) - 1
				cStruct(index).sName = CStr(reader.Item("sapstruct"))
				cStruct(index).sData = CStr(reader.Item("datastruct"))
				cStruct(index).sIndex = index
				cStruct(index).sLength = CInt(reader.Item("length"))
				cStruct(index).sLevel = CInt(reader.Item("slevel"))
				cStruct(index).sType = CChar(reader.Item("datatype"))
				cStruct(index).sFunction = CStr(reader.Item("rfcfunction"))
			End While
		End If
		readDBStructures = True
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
#If Not MitaOrder And Not MitaError And Not DLL Then
		MsgBox(mitaData.errorDescription)
		End
#End If
		readDBStructures = True
		Resume Exx
	End Function
	Public Function readDBEvents(ByRef cEvents() As EVENTSTRUCT, ByRef cCount As Integer) As Boolean
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		Dim query As String
		Dim a As String
		readDBEvents = False
		On Error GoTo isErr
		Dim sb As New System.Text.StringBuilder(1000)

		sb.Append("SELECT DISTINCT event, comments FROM " & mitaSystem.tableEventControl)
		sb.Append(" WHERE sapsystemid = " & mitaSystem.sapSystemID)
		sb.Append(" AND rfctype = " & mitaSystem.rfcType)
		sb.Append(" AND action IS NULL")
		sb.Append(" AND activ = 'Y'")
		sb.Append(" ORDER BY event;")
		query = sb.ToString()
		cCount = -1
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			cCount = cCount + 1
			ReDim Preserve cEvents(cCount)
			cEvents(cCount).sName = CStr(reader.Item("event"))
			cEvents(cCount).sComment = CStr(reader.Item("comments"))
			cEvents(cCount).sCount = -1
		End While
		readDBEvents = True
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
#If Not MitaOrder And Not MitaError And Not DLL Then
		MsgBox(mitaData.errorDescription)
		End
#End If
		Resume Exx
	End Function

	Public Function readDBEventRunTypes(ByRef cTypes() As String) As Boolean
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		Dim query As String
		Dim a As String
		Dim cCount As Integer
		readDBEventRunTypes = False
		On Error GoTo isErr
		Dim sb As New System.Text.StringBuilder(1000)

		sb.Append("SELECT DISTINCT runtype FROM " & mitaSystem.tableEventControl)
		sb.Append(" WHERE sapsystemid = " & mitaSystem.sapSystemID)
		sb.Append(" AND action IS NULL")
		sb.Append(" AND activ = 'Y'")
		sb.Append(" ORDER BY runtype;")
		query = sb.ToString()
		cCount = -1
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			cCount = cCount + 1
			ReDim Preserve cTypes(cCount)
			cTypes(cCount) = CStr(reader.Item("runtype"))
		End While
		readDBEventRunTypes = True
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
#If Not MitaOrder And Not MitaError And Not DLL Then
		MsgBox(mitaData.errorDescription)
		End
#End If
		Resume Exx
	End Function

	Public Sub readDBTablePairs(ByRef Index As Integer)
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		Dim query As String
		Dim a As String
		Dim name As String
		Dim tmptable As TABLESTRUCT = mitaData.tables(Index)
		Dim sb As New System.Text.StringBuilder(1000)
		On Error GoTo isErr
		sb.Append("SELECT tablename, tleft, tright FROM " & mitaSystem.tableCustTables)
		sb.Append(" WHERE activ = 'Y'")
		sb.Append(" AND sapsystemid = " & mitaSystem.sapSystemID)
		sb.Append(" AND rfctype = " & mitaSystem.rfcType)
		sb.Append(" AND tablename = '" & tmptable.tName & "'")
		sb.Append(" AND tleft = 'DEFAULT';")
		query = sb.ToString()
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		tmptable.tCount = -1
		While reader.Read
			tmptable.tCount = tmptable.tCount + 1
			ReDim Preserve tmptable.tEntries(tmptable.tCount)
			Dim tmp As String
			If Not IsDBNull(reader.Item("tleft")) Then
				tmp = CStr(reader.Item("tleft"))
			Else
				tmp = ""
			End If
			tmptable.tEntries(tmptable.tCount).pLeft = tmp
			If Not IsDBNull(reader.Item("tright")) Then
				tmp = CStr(reader.Item("tright"))
			Else
				tmp = ""
			End If
			tmptable.tEntries(tmptable.tCount).pRight = tmp
		End While
		If Not IsNothing(reader) Then reader.Close()
		sb.Remove(0, sb.Length)
		sb.Append("SELECT tablename, tleft, tright FROM " & mitaSystem.tableCustTables)
		sb.Append(" WHERE activ = 'Y'")
		sb.Append(" AND tablename = '" & tmptable.tName & "'")
		sb.Append(" AND tleft <> 'DEFAULT' ORDER BY tleft;")
		query = sb.ToString()
		idbc.CommandText = query
		reader = idbc.ExecuteReader()
		While reader.Read
			tmptable.tCount = tmptable.tCount + 1
			ReDim Preserve tmptable.tEntries(tmptable.tCount)
			Dim tmp As String
			If Not IsDBNull(reader.Item("tleft")) Then
				tmp = CStr(reader.Item("tleft"))
			Else
				tmp = ""
			End If
			tmptable.tEntries(tmptable.tCount).pLeft = tmp
			If Not IsDBNull(reader.Item("tright")) Then
				tmp = CStr(reader.Item("tright"))
			Else
				tmp = ""
			End If
			tmptable.tEntries(tmptable.tCount).pRight = tmp
		End While
		mitaData.tables(Index) = tmptable
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Sub
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
#If Not MitaOrder And Not MitaError And Not DLL Then
		MsgBox(mitaData.errorDescription)
		End
#End If
		Resume Exx
	End Sub

	Public Function readDBEventActions(ByRef eName As String, ByRef cEvents() As EVENTSTRUCT, ByRef cIndex As Short) As Boolean
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		Dim query As String
		Dim a As String
		Dim cnt As Integer
		readDBEventActions = False
		Dim sb As New System.Text.StringBuilder(1000)

		sb.Append("SELECT action, recno, parameter, comments FROM " & mitaSystem.tableEventControl)
		sb.Append(" WHERE sapsystemid = " & mitaSystem.sapSystemID)
		sb.Append(" AND rfctype = " & mitaSystem.rfcType)
		sb.Append(" AND runtype = '" & mitaSystem.runType & "'")
		sb.Append(" AND event = '" & eName & "'")
		sb.Append(" AND action IS NOT NULL")
		sb.Append(" AND activ = 'Y'")
		sb.Append(" ORDER BY recno;")
		query = sb.ToString()
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		cnt = -1
		While reader.Read
			cnt = cnt + 1
			ReDim Preserve cEvents(cIndex).sActions(cnt)
			cEvents(cIndex).sActions(cnt).aName = CStr(reader.Item("action"))
			If IsDBNull(reader.Item("comments")) Then
				cEvents(cIndex).sActions(cnt).aComment = ""
			Else
				cEvents(cIndex).sActions(cnt).aComment = CStr(reader.Item("comments"))
			End If
			cEvents(cIndex).sActions(cnt).aRecno = CInt(reader.Item("recno"))
			If IsDBNull(reader.Item("parameter")) Then
				cEvents(cIndex).sActions(cnt).aTodo = ""
			Else
				cEvents(cIndex).sActions(cnt).aTodo = Replace(CStr(reader.Item("parameter")), "~", "'")
			End If
		End While
		cEvents(cIndex).sCount = cnt
		readDBEventActions = True
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
#If Not MitaOrder And Not MitaError And Not DLL Then
		MsgBox(mitaData.errorDescription)
		End
#End If
		Resume Exx
	End Function



	Public Function readDBStructFields(ByRef sName As String, ByRef cFields() As FIELDSTRUCT, ByRef cCount As Integer) As Boolean
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		Dim query As String
		readDBStructFields = False
		Dim sb As New System.Text.StringBuilder(1000)
		sb.Append("SELECT field, recno, first, length from " & mitaSystem.tableStructFields)
		sb.Append(" WHERE sapstruct = '" & sName & "'")
		sb.Append(" AND sapversionid = " & mitaSystem.sapSystemVERSIONID)
		sb.Append(" AND rfctype = " & mitaSystem.rfcType)
		sb.Append(" AND activ = 'Y'")
		sb.Append(" ORDER BY recno;")
		query = sb.ToString()
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		ReDim cFields(0)
		cCount = -1
		While reader.Read
			cCount = cCount + 1
			ReDim Preserve cFields(cCount)
			cFields(cCount).fName = UCase(CStr(reader.Item("field")))
			cFields(cCount).fFirst = CInt(reader.Item("first"))
			cFields(cCount).fLength = CInt(reader.Item("length"))
			cFields(cCount).fIndex = CInt(reader.Item("recno"))
		End While
		readDBStructFields = True
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
#If Not MitaOrder And Not MitaError And Not DLL Then
		MsgBox(mitaData.errorDescription)
		End
#End If
		Resume Exx
	End Function
	Public Function readDBReplaces(ByRef repl() As String, ByVal versionID As Integer) As Boolean
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		Dim query As String
		readDBReplaces = False
		Dim tmp As String
		Dim x() As String
		Dim cnt As Integer = -1
		Dim i As Integer
		ReDim repl(0)
		On Error GoTo isErr
		Dim sb As New System.Text.StringBuilder(1000)

		sb.Append("SELECT replacestruct, replacefunc FROM " & mitaSystem.tableSapVersions)
		sb.Append(" WHERE sapversionid = " & versionID)
		sb.Append(" AND activ = 'Y'")
		query = sb.ToString()
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		If reader.Read Then
			If Not IsDBNull(reader.Item("replacestruct")) Then
				tmp = CStr(reader.Item("replacestruct"))
				x = Split(tmp, ",")
				For i = 0 To UBound(x)
					cnt = cnt + 1
					ReDim Preserve repl(cnt)
					repl(cnt) = x(i)
				Next
				readDBReplaces = True
			End If
			If Not IsDBNull(reader.Item("replacefunc")) Then
				tmp = CStr(reader.Item("replacefunc"))
				x = Split(tmp, ",")
				For i = 0 To UBound(x)
					cnt = cnt + 1
					ReDim Preserve repl(cnt)
					repl(cnt) = x(i)
				Next
				readDBReplaces = True
			End If
		End If
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
		MsgBox(mitaData.errorDescription)
	End Function

	Public Function getHostList(ByRef cHosts() As String, ByRef cCount As Integer) As Boolean
		Dim query As String
		Dim sb As New System.Text.StringBuilder(1000)
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		getHostList = False
		cCount = -1
		sb.Append("SELECT DISTINCT loginhost from " & mitaSystem.tableOnline)
		sb.Append(" ORDER BY loginhost ASC")
		query = sb.ToString()
		idbc.CommandText = query
		On Error GoTo isErr
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			cCount = cCount + 1
			ReDim Preserve cHosts(cCount)
			cHosts(cCount) = CStr(reader.Item("loginhost"))
		End While
		getHostList = True
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query
		Resume Exx
	End Function
End Module