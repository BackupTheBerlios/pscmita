Option Strict Off
Option Explicit On 
Imports pscMitaDef.CMitaDef

Module modMitaOLPSQL

	Public appNumberAd As String
	Public appNumberVersion As Integer
	Public appNumberRefAd As String
	Public appCountInsertion As Integer
	Public appStartDate As String
	Public appEndDate As String
	Public appYSize As Integer
	Public appXSize As Integer
	Public appXLoc As Integer
	Public appYLoc As Integer
	Public appWidthAd As Integer
	Public appStatusWord As String
	Public appText As String
	Public appDesignFileName As String
	Public numberInsertion As Integer

	Public SapPool As pscPoolStruct.CSapPool
	Public SapPoolIn As pscPoolStruct.CSapPool

	Public AllSQLThirdParty() As SQLLOG
	Public AllSQLSAP() As SQLLOG

	Private Function getNewVer() As Integer
		Dim i As Integer
		Dim sSql As SQLLOG
		Dim ok As Boolean
		getNewVer = 0
		If appNumberAd = "" Then Exit Function
		colorDB()
		sSql = sqlBuild(mitaSqlClass.classAd, "GetNewAdVer")
		SapOrder.eventLog(sSql)
		mitaConnect.queryNumber(sSql.sResult, getNewVer)
		colorRestore()
	End Function

	Private Function newAdNo() As Integer
		Dim sSql As SQLLOG
		newAdNo = 0
		colorDB()
		sSql.sName = "NewAdNo"
		sSql.sClass = mitaSqlClass.classInternal
		If SapOrder.getCustomerQuery(sSql) Then
			SapOrder.eventLog(sSql)
			If Not mitaConnect.queryNumber(sSql.sResult, newAdNo) Then
				If SapOrder.eventRaise(sSql.sResult, mitaEventCodes.errorDatabaseSequence, "newAdNo") Then
					Debugger.Break()
				End If
			End If
		End If
		colorRestore()
	End Function

	'Private Function newContentNo() As Integer
	'	Dim sSql As SQLLOG
	'	newContentNo = 0
	'	colorDB()
	'	sSql.sName = "NewContentNo"
	'	sSql.sClass = mitaSqlClass.classInternal
	'	If SapOrder.getCustomerQuery(sSql) Then
	'		SapOrder.eventLog(sSql)
	'		If Not sqlQueryNumber(sSql.sResult, newContentNo) Then
	'			If SapOrder.eventRaise(sSql.sResult, (mitaEventCodes.errorDatabaseSequence), "newContentNo") Then
	'				Debugger.Break()
	'			End If
	'		End If
	'		colorRestore()
	'	End If
	'End Function

	Public Sub sapToDB()
		Dim orderVerNo As Integer
		Dim sapCnt As Integer
		Dim dt As String

		ReDim AllSQLSAP(0)
		ReDim AllSQLThirdParty(0)
		abortOrder = False

		If doProfile Then myProfile.profileStart("sapToDB")
		orderVerNo = SapOrder.valueOrderVno
		If orderVerNo < getspsOrderVerNo() Then
			SapOrder.eventRaise("", (mitaEventCodes.errorSAPVersion))
			SapOrder.eventRaise("", mitaEventCodes.userOrderFailure)
			Exit Sub
		End If
		Do
			If Not SapOrder.orderNextMotiv Then Exit Do
			If Not chkDontUse() Then
				odbc_connection.Open()
				Do
					If Not SapOrder.orderNextPub Then Exit Do
					appNumberAd = ""
					appNumberVersion = 0
					appEndDate = "00000000"
					appStartDate = "99999999"

					Do
						If Not SapOrder.orderNextCombo Then Exit Do
						sapCnt = readSapPool(SapPool)
						If Not IsNothing(SapPoolIn) Then SapPoolIn.Dispose()
						SapPoolIn = SapPool.clone
						SapOrder.infoSetAdNo(0)
						SapOrder.infoSetAdVer(0)
						appCountInsertion = 0
						Do
							If Not SapOrder.orderNextInsertion Then Exit Do
							Dim DelFlg As Boolean = (Val(getSapField("DELETESTATUS")) = 1)
							If DelFlg Then
								appStatusWord = "PER"
							Else
								appStatusWord = "VAR"
							End If

							dt = CStr(Val(getSapField("RUNDATE")))
							If dt > appEndDate Then appEndDate = dt
							If dt < appStartDate Then appStartDate = dt

							appCountInsertion = appCountInsertion + 1
							appNumberAd = CStr(SapPool.spsAdNo)
							If appCountInsertion = 1 And Val(appNumberAd) <= 0 Then
								appNumberAd = CStr(newAdNo())
								appNumberVersion = 1
							ElseIf appCountInsertion = 1 Then
								appNumberVersion = getNewVer() + 1
							End If
							If appCountInsertion = 1 Then
								SapOrder.infoSetAdNo(CInt(appNumberAd))
								SapOrder.infoSetAdVer(appNumberVersion)
								sqlBuildAd("SetOldVersion")
								If sapCnt = 0 Then
									SapPool.spsAvm = SapOrder.valueOrderString
									SapPool.spsMotivno = SapOrder.valueMotivString
									SapPool.spsComboNo = SapOrder.valueComboNo
									SapPool.spsAdType = getSapField("ADTYPE")
									SapPool.spsClientNo = getSapField("CLIENTNO")
								End If
								If SapPool.spsDsnNam = "" Then
									SapPool.spsDsnNam = getSapField("ADVNAME")
								End If
								SapPool.spsPosNo = SapOrder.valuePosNo
								appDesignFileName = SapPool.spsDsnNam
								If SapPool.spsAdSort = "" Then
									SapPool.spsAdSort = Trim(getSapField("SORTWORD"))
								End If
								appText = stringForDb(SapPool.spsAdSort)
								appXLoc = 0
								appXLoc = 0
								appXSize = 0
								appYSize = CInt(0.1 * Int(10 * CInt(getSapField("DEPTH")) + 0.5))
								SapPool.spsYSize = appYSize
								SapPool.spsXSize = appXSize
								SapPool.spsXLoc = appXLoc
								SapPool.spsYLoc = appYLoc
								SapPool.spsBoxNo = getSapField("BOXNO")
								SapPool.spsAdNo = CInt(appNumberAd)
								SapPool.spsVno = orderVerNo
							End If
							If abortOrder Then Exit Do
							If Not SapPool.equals(SapPoolIn) Then
								sqlBuildSAP("UpdateSapPool")
								SapPoolIn = SapPool.clone
							End If
							sqlBuildAd("InsertPub")
						Loop
						sqlBuildAd("InsertAd")
						If appEndDate <> "00000000" Then
							sqlBuildAd("UpdateAdFinally")
						End If
						If Not SapOrder.itemTestLastCombo Then
							If abortOrder Then Exit Do
							sqlExecAll(AllSQLThirdParty)
							If abortOrder Then Exit Do
							sqlExecAll(AllSQLSAP)
							If abortOrder Then Exit Do
							SapOrder.eventRaise("", (mitaEventCodes.userMotivSuccess))
						End If
					Loop
				Loop
				If SapOrder.itemTestReference() Then storeReference()
				If abortOrder Then Exit Do
				sqlExecAll(AllSQLThirdParty)
				If abortOrder Then Exit Do
				sqlExecAll(AllSQLSAP)
				If Not abortOrder Then SapOrder.eventRaise("", (mitaEventCodes.userMotivSuccess))
				odbc_connection.Close()
			End If
			If abortOrder Then Exit Do
		Loop
		If odbc_connection.State = ConnectionState.Open Then odbc_connection.Close()
		If Not abortOrder Then
			SapOrder.eventRaise("", mitaEventCodes.userOrderSuccess)
		Else
			SapOrder.eventRaise("", mitaEventCodes.userOrderFailure)
		End If
		If doProfile Then myProfile.profileEnd("sapToDB")
	End Sub

	Private Function getSpsOrderVerNo() As Integer
		Dim sSql As SQLLOG
		sSql = sqlBuild(mitaSqlClass.classOrder, "GetSapVersion")
		SapOrder.eventLog(sSql)
		getspsOrderVerNo = -1
		odbc_connection.Open()
		mitaConnect.queryNumber(sSql.sResult, getSpsOrderVerNo)
		odbc_connection.Close()
	End Function

	Private Sub sqlBuildAd(ByRef SrcSQL As String)
		Dim i As Integer
		Dim NewSQL As SQLLOG
		NewSQL = sqlBuild(mitaSqlClass.classAd, SrcSQL)
		i = UBound(AllSQLThirdParty) + 1
		If AllSQLThirdParty(0).sResult = "" And i = 1 Then i = 0
		ReDim Preserve AllSQLThirdParty(i)
		If NewSQL.sNumberAd = "" Then NewSQL.sNumberAd = appNumberAd
		AllSQLThirdParty(i) = NewSQL
	End Sub

	Private Sub sqlBuildSAP(ByRef SrcSQL As String)
		Dim i As Integer
		Dim NewSQL As SQLLOG
		NewSQL = sqlBuild(mitaSqlClass.classOrder, SrcSQL)
		i = UBound(AllSQLSAP) + 1
		If AllSQLSAP(0).sResult = "" And i = 1 Then i = 0
		ReDim Preserve AllSQLSAP(i)
		If appNumberAd <> "" Then NewSQL.sNumberAd = appNumberAd
		If NewSQL.sNumberAd = "" Then NewSQL.sNumberAd = appNumberAd
		AllSQLSAP(i) = NewSQL
	End Sub

	Private Sub sqlExecAll(ByRef source() As SQLLOG)
		Dim i As Integer
		Dim allOk As Boolean
		Dim t As String
		Dim f As Integer
		Dim idbc As OdbcCommand = odbc_connection.CreateCommand
		Dim retCode As mitaEventReturnCodes
		colorDB()
		If Not abortOrder Then
			On Error GoTo IsErr
			oracleTrans = odbc_connection.BeginTransaction
			idbc.Transaction = oracleTrans
			allOk = True
			For i = 0 To UBound(source)
				If source(i).sResult > "" Then
					If doProfile Then myProfile.profileStart("execAllSQl, " & source(i).sName)
					idbc.CommandText = source(i).sResult
					idbc.ExecuteNonQuery()
					If Not allOk Then Exit For
					SapOrder.eventLog(source(i))
					If doProfile Then myProfile.profileEnd("execAllSQl, " & source(i).sName)
				End If
			Next i
			If allOk Then
				oracleTrans.Commit()
			Else
				oracleTrans.Rollback()
			End If
		End If
exx:
		ReDim source(0)
		colorRestore()
		Exit Sub
IsErr:
		allOk = False
		source(i).sError = Err.Description
		SapOrder.eventLog(source(i))
		If doProfile Then myProfile.profileEnd("execAllSQl, " & source(i).sName)
		retCode = SapOrder.eventRaise(Err.Description & vbCrLf & idbc.CommandText, mitaEventCodes.userSqlException, "sqlExecAll")
		If ((retCode And mitaEventReturnCodes.debugBreak) = mitaEventReturnCodes.debugBreak) Then
			Debugger.Break()
		End If
		If ((retCode And mitaEventReturnCodes.requeryExceeded) = mitaEventReturnCodes.requeryExceeded) Then
			oracleTrans.Rollback()
			Resume exx
		ElseIf ((retCode And mitaEventReturnCodes.requeryRequest) = mitaEventReturnCodes.requeryRequest) Then
			allOk = True
			Resume
		End If
	End Sub

	Private Function readSapPool(ByRef trg As pscPoolStruct.CSapPool) As Integer
		Dim sSql As SQLLOG
		Dim sapcnt As Integer
		Dim dbRes As DataSet = New DataSet
		Dim rw As DataRow
		Dim cnt As Integer
		If Not IsNothing(trg) Then trg.Dispose()
		trg = New pscPoolStruct.CSapPool
		sSql = sqlBuild(mitaSqlClass.classOrder, "ReadSapPool")
		sapcnt = 0
		If sqlExecQuery(sSql, dbRes) Then
			For cnt = 0 To dbRes.Tables(0).Rows.Count - 1
				sapcnt = sapcnt + 1
				rw = dbRes.Tables(0).Rows(sapcnt - 1)
				If Not IsDBNull(rw.Item("adno")) Then
					trg.spsAdNo = CType(rw.Item("adno"), Integer)
				End If
				If Not IsDBNull(rw.Item("avm")) Then
					trg.spsAvm = CType(rw.Item("avm"), String)
				End If
				If Not IsDBNull(rw.Item("ordervno")) Then
					trg.spsVno = CType(rw.Item("ordervno"), Integer)
				End If
				If Not IsDBNull(rw.Item("motivno")) Then
					trg.spsMotivno = CType(rw.Item("motivno"), Integer)
				End If
				If Not IsDBNull(rw.Item("dsnfile")) Then
					trg.spsDsnNam = CType(rw.Item("dsnfile"), String)
				End If
				If Not IsDBNull(rw.Item("txtfile")) Then
					trg.spsTxtNam = CType(rw.Item("txtfile"), String)
				End If
				If Not IsDBNull(rw.Item("combono")) Then
					trg.spsComboNo = CType(rw.Item("combono"), Integer)
				End If
				If Not IsDBNull(rw.Item("posno")) Then
					trg.spsPosNo = CType(rw.Item("posno"), Integer)
				End If
				If Not IsDBNull(rw.Item("adtype")) Then
					trg.spsAdType = CType(rw.Item("adtype"), String)
				End If
				If Not IsDBNull(rw.Item("boxno")) Then
					trg.spsBoxNo = CType(rw.Item("boxno"), String)
				End If
				If Not IsDBNull(rw.Item("clientno")) Then
					trg.spsClientNo = CType(rw.Item("clientno"), String)
				End If
				If Not IsDBNull(rw.Item("sortword")) Then
					trg.spsAdSort = CType(rw.Item("sortword"), String)
				End If
				If Not IsDBNull(rw.Item("ysize")) Then
					trg.spsYSize = CType(rw.Item("ysize"), Integer)
				End If
				If Not IsDBNull(rw.Item("xsize")) Then
					trg.spsXSize = CType(rw.Item("xsize"), Integer)
				End If
				If Not IsDBNull(rw.Item("xloc")) Then
					trg.spsXLoc = CType(rw.Item("xloc"), Integer)
				End If
				If Not IsDBNull(rw.Item("yloc")) Then
					trg.spsYLoc = CType(rw.Item("yloc"), Integer)
				End If
			Next cnt
		End If
		dbRes.Dispose()
		Return (sapcnt)
	End Function
	Private Function chkDontUse() As Boolean
		Dim result As Boolean
		Dim tmp As String
		chkDontUse = False
		result = SapOrder.getStatusValue(7, "DONTUSE", tmp)
		If result Then
			If tmp <> "" Then
				chkDontUse = True
			End If
		End If
	End Function

	Private Sub storeReference()
		Dim sSql As SQLLOG
		If doProfile Then myProfile.profileStart("storeReference")
		sSql = sqlBuild(mitaSqlClass.classAd, "GetReferenceAdNo")
		If mitaConnect.queryNumber(sSql.sResult, appNumberRefAd) Then
			sqlBuildAd("WriteReference")
		Else
			sSql.sError = mitaConnect.errorDescription
			SapOrder.eventRaise("", CInt(mitaEventCodes.errorSAPReference), "storeReference")
		End If
		SapOrder.eventLog(sSql)
		If doProfile Then myProfile.profileEnd("storeReference")
	End Sub
End Module