Imports pscMitaDef.CMitaDef
Module modMitaCommonSQL
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
		result.sNumberAd = appNumberAd
		result.sVersionAd = appNumberVersion
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
					NewSQL.Append(CStr(Val(appNumberAd)))
				ElseIf tmp = "APPADV" Then
					NewSQL.Append(stringForDb(appDesignFileName))
				ElseIf tmp = "APPCREDP" Then
					SapOrder.getStatusValue(8, "MERKMAL0", t)
					NewSQL.Append(stringForDb(t))
				ElseIf tmp = "APPDEPTH" Then
					NewSQL.Append(CStr(appYSize))
				ElseIf tmp = "APPWIDTH" Then
					NewSQL.Append(CStr(appXSize))
				ElseIf tmp = "APPENDDATE" Then
					NewSQL.Append(appEndDate)
				ElseIf tmp = "APPFP" Then
					If appXLoc > 0 Then
						NewSQL.Append("1")
					Else
						NewSQL.Append("0")
					End If
				ElseIf tmp = "APPREFADNO" Then
					NewSQL.Append(CStr(Val(appNumberRefAd)))
				ElseIf tmp = "APPSTARTDATE" Then
					NewSQL.Append(appStartDate)
				ElseIf tmp = "APPSTATUS" Then
					NewSQL.Append(appStatusWord)
				ElseIf tmp = "APPTEXT" Then
					NewSQL.Append(appText)
				ElseIf tmp = "APPVNO" Then
					NewSQL.Append(CStr(appNumberVersion))
				ElseIf tmp = "APPXLOC" Then
					NewSQL.Append(CStr(appXLoc))
				ElseIf tmp = "APPYLOC" Then
					NewSQL.Append(CStr(appYLoc))
				Else
					MsgBox("Not found: " & tmp)
				End If
			End If
		Loop
		result.sResult = NewSQL.ToString
		If doProfile Then myProfile.profileEnd("sqlBuild, " & SrcSQL)
		sqlBuild = result
	End Function
	Public Function sqlExecQuery(ByRef statement As SQLLOG, ByRef dbRes As DataSet) As Boolean
		sqlExecQuery = False
		colorDB()
		If doProfile Then myProfile.profileStart("sqlExecQuery, " & statement.sName)
		Try
			Dim da As OdbcDataAdapter = New OdbcDataAdapter(statement.sResult, odbc_connection)
			da.Fill(dbRes)
			sqlExecQuery = True
		Catch
			statement.sError = Err.Description
			If SapOrder.eventRaise(Err.Description, (mitaEventCodes.userSqlException), "sqlExecQuery") Then
				Debugger.Break()
			End If
		End Try
		SapOrder.eventLog(statement)
		colorRestore()
		If doProfile Then myProfile.profileEnd("sqlExecQuery, " & statement.sName)
	End Function
End Module
