Imports MitaDefClass.MitaDef
Module modMitaCommonSQL
	Public Function stringForDb(ByRef buf1 As String) As String
		Dim buf2 As String
		buf2 = Replace(buf1, "'", " ")
		buf2 = Replace(buf2, "@", " ")
		buf2 = Replace(buf2, "´", " ")
		buf2 = Replace(buf2, "`", " ")
		stringForDb = buf2
	End Function
	Public Function sqlBuild(ByVal sClass As mitaSqlClass, ByRef SrcSQL As String) As SQLLOG
		Dim iBeg, iPos As Integer
		Dim searchStart As Integer
		Dim tmp As String
		Dim NewSQL As New System.text.StringBuilder(1000)
		Dim Sql As String
		Dim ok As Boolean
		Dim buffer As String
		Dim t As String
		Dim result As SQLLOG
		result.sName = SrcSQL
		result.sClass = sClass
		result.sNumberAd = numberAd
		result.sVersionAd = numberVersion
		If doProfile Then myProfile.profileStart("sqlBuild, " & SrcSQL)
		ok = SapOrder.getCustomerQuery(result)
		Sql = result.sResult
		searchStart = 1
		Do
			iPos = InStr(searchStart, Sql, "#")
			If iPos = 0 Then
				NewSQL.Append(Mid$(Sql, searchStart))
				Exit Do
			End If
			NewSQL.Append(Mid(Sql, searchStart, iPos - searchStart))
			searchStart = iPos + 1

			iPos = InStr(searchStart, Sql, "#")
			If iPos = 0 Then
				NewSQL.Append("#" & Mid$(Sql, searchStart))
				Exit Do
			End If
			tmp = UCase(Mid$(Sql, searchStart, iPos - searchStart))
			searchStart = iPos + 1
			If Left(tmp, 3) = "APP" Then
				If tmp = "APPADNO" Then
					NewSQL.Append(CStr(Val(numberAd)))
				ElseIf tmp = "APPADV" Then
					NewSQL.Append(stringForDb(Trim(Sap2A.S2AFilNam)))
				ElseIf tmp = "APPCOMBONO" Then
					NewSQL.Append(stringForDb(Trim(Sap2A.S2AComboNo)))
				ElseIf tmp = "APPCREDP" Then
					SapOrder.getStatusValue(8, "MERKMAL0", t)
					NewSQL.Append(stringForDb(t))
				ElseIf tmp = "APPDEPTH" Then
					NewSQL.Append(CStr(Sap2A.S2AYSizeMM))
				ElseIf tmp = "APPENDDATE" Then
					NewSQL.Append(EndDate)
				ElseIf tmp = "APPFP" Then
					If intXLoc > 0 Then
						NewSQL.Append("1")
					Else
						NewSQL.Append("0")
					End If
				ElseIf tmp = "APPPUBCNT" Then
					NewSQL.Append(CStr(countInsertion))
				ElseIf tmp = "APPREFADNO" Then
					NewSQL.Append(CStr(Val(numberReferenceAd)))
				ElseIf tmp = "APPREFAVMNO" Then
					NewSQL.Append(numberReferenceAvm)
				ElseIf tmp = "APPREFPOSNO" Then
					NewSQL.Append(CStr(Val(numberReferencePos)))
				ElseIf tmp = "APPSTARTDATE" Then
					NewSQL.Append(StartDate)
				ElseIf tmp = "APPSTATUS" Then
					NewSQL.Append(statusWord)
				ElseIf tmp = "APPTEXT" Then
					t = stringForDb(Sap2A.S2AAdSort)
					If t = "" Then t = stringForDb(getSapField("SORTWORD"))
					NewSQL.Append(t)
				ElseIf tmp = "APPVNO" Then
					NewSQL.Append(CStr(numberVersion))
				ElseIf tmp = "APPXLOC" Then
					NewSQL.Append(CStr(Sap2A.S2AXLoc))
				ElseIf tmp = "APPYLOC" Then
					NewSQL.Append(CStr(Sap2A.S2AYLoc))
				Else
					MsgBox("Not found: " & tmp)
				End If
			End If
		Loop
		result.sResult = NewSQL.ToString
		If doProfile Then myProfile.profileEnd("sqlBuild, " & SrcSQL)
		sqlBuild = result
	End Function
	Public Function sqlBuildAndExecute(ByVal sClass As mitaSqlClass, ByRef SrcSQL As String) As SQLLOG
		If Not abortOrder Then
			Dim sql As SQLLOG = sqlBuild(sClass, SrcSQL)
			Dim idbc As OdbcCommand = odbc_connection.CreateCommand
			colorDB()
			Try
				oracleTrans = odbc_connection.BeginTransaction
				idbc.Transaction = oracleTrans
				If doProfile Then myProfile.profileStart("sqlBuildAndExecute, " & sql.sName)
				SapOrder.eventLog(sql)
				idbc.CommandText = sql.sResult
				idbc.ExecuteNonQuery()
				If doProfile Then myProfile.profileEnd("sqlBuildAndExecute, " & sql.sName)
				oracleTrans.Commit()
				idbc.Dispose()
			Catch
				oracleTrans.Rollback()
				If SapOrder.eventRaise(Err.Description, (mitaEventCodes.userSqlException), "sqlExecAll") Then
					Debugger.Break()
				End If
			End Try
			colorRestore()
		End If
	End Function
	Public Function sqlQueryNumber(ByRef query As String, ByRef number As Integer) As Boolean
		Dim idbc As OdbcCommand = odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		sqlQueryNumber = False
		idbc.CommandText = query
		Try
			reader = idbc.ExecuteReader()
			sqlQueryNumber = False
			If reader.Read Then
				number = reader.GetValue(0)
				sqlQueryNumber = True
			End If
			reader.Close()
			idbc.Dispose()
		Catch
			If SapOrder.eventRaise(Err.Description, (mitaEventCodes.userSqlException), "sqlQueryNumber") Then
				Debugger.Break()
			End If
		End Try
	End Function
	Public Function sqlQueryString(ByRef query As String, ByRef target As String) As Boolean
		Dim idbc As OdbcCommand = odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		sqlQueryString = False
		idbc.CommandText = query
		Try
			reader = idbc.ExecuteReader()
			sqlQueryString = False
			If reader.Read Then
				target = reader.GetString(0)
				sqlQueryString = True
			End If
			reader.Close()
		Catch
			If SapOrder.eventRaise(Err.Description, (mitaEventCodes.userSqlException), "sqlQueryString") Then
				Debugger.Break()
			End If
		Finally
			idbc.Dispose()
		End Try
	End Function
	Public Function sqlExecQuery(ByRef statement As SQLLOG, ByRef dbRes As DataSet) As Boolean
		sqlExecQuery = False
		colorDB()
		sqlExecQuery = False
		If doProfile Then myProfile.profileStart("sqlExecQuery, " & statement.sName)
		SapOrder.eventLog(statement)
		Try
			Dim da As OdbcDataAdapter = New OdbcDataAdapter(statement.sResult, odbc_connection)
			da.Fill(dbRes)
			sqlExecQuery = True
		Catch
			If SapOrder.eventRaise(Err.Description, (mitaEventCodes.userSqlException), "sqlExecQuery") Then
				Debugger.Break()
			End If
		Finally
			colorRestore()
			If doProfile Then myProfile.profileEnd("sqlExecQuery, " & statement.sName)
		End Try
	End Function
	Public Function getSapField(ByRef fieldName As String) As String
		Dim result As Boolean
		Dim tmp As String
		result = SapOrder.getFieldValue(fieldName, tmp)
		If result Then
			getSapField = tmp
		Else
			getSapField = ""
		End If
	End Function
End Module
