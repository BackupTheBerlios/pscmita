Option Strict Off
Option Explicit On 
Imports VB = Microsoft.VisualBasic
Imports System.Data.Odbc
Imports pscSapServer.CSapServer
Imports pscMitaDef.CMitaDef
Public Class CMitaOrder

	Private Const hashCodeCount As Integer = 1023


	Private Const cUpdateControl As String = "UPDATE #CONTROL# SET activ = 'N',comments = TO_CHAR(sysdate,'DD.MM.YYYY HH24:MI:SS') WHERE avm = '#AVM#' AND version = #VNO# AND sapversionid = #SAPVERSION# AND sapsystemid = #SAPSYSTEM# AND pscid <> #ORDERID# AND activ = 'Y'"
	Private Const cRetryControl As String = "UPDATE #CONTROL# SET activ = 'N',comments = '#LASTERROR#' WHERE avm = '#AVM#' AND version = #VNO# AND sapversionid = #SAPVERSION# AND sapsystemid = #SAPSYSTEM# AND pscid <> #ORDERID#"
	Private Const cFind As String = "SELECT p1.length, p2.structures, p1.avm, p1.version FROM #CONTROL# p1, #DATA# p2 WHERE p1.pscid = #ORDERID# AND p2.pscid = p1.pscid AND p1.sapversionid = #SAPVERSION# AND p1.sapsystemid = #SAPSYSTEM#"
	Private Const cGetID As String = "SELECT pscid from #CONTROL# WHERE avm = '#AVM#' AND sapversionid = #SAPVERSION# AND sapsystemid = #SAPSYSTEM#"
	Private Const cGetAvm As String = "SELECT avm FROM #CONTROL# WHERE pscid = #ORDERID# AND sapversionid = #SAPVERSION# AND sapsystemid = #SAPSYSTEM#"
	Private Const cGetVno As String = "SELECT version FROM #CONTROL# WHERE pscid = #ORDERID# AND sapversionid = #SAPVERSION# AND sapsystemid = #SAPSYSTEM#"
	Private Const cGetByVersion As String = "SELECT pscid from #CONTROL# WHERE avm = '#AVM#' AND sapversionid = #SAPVERSION# AND sapsystemid = #SAPSYSTEM# and version = #VNO#"
	Private Const cGetVersion As String = "SELECT version from #CONTROL# WHERE avm = '#AVM#' AND sapversionid = #SAPVERSION# AND sapsystemid = #SAPSYSTEM# ORDER BY version DESC"
	Private Const cNewData As String = "INSERT INTO #DATA# VALUES(#ORDERID#, NULL, NULL)"
	Private Const cNewControl As String = "INSERT INTO #CONTROL# (pscid, avm, version, createid, createhost, status, length, sapversionid, sapsystemid, processed) VALUES(#ORDERID#, '#AVM#', #VNO#, #MYID#, '#HOST#', 'N', #LENGTH#, #SAPVERSION#, #SAPSYSTEM#, #PROCESSED#)"
	Private Const cWriteBlobs As String = "UPDATE #DATA# SET hashcode = ?, structures = ? WHERE pscid = #ORDERID#"
	Private Const cDeactivate As String = "UPDATE #CONTROL# SET activ = 'N' WHERE avm = '#AVM#' AND version < #VNO# AND sapversionid = #SAPVERSION# AND sapsystemid = #SAPSYSTEM#"
	Private Const cNewLog As String = "INSERT INTO #LOG# (pscid, createid, createhost, createuser, priority, text, sapsystemid, rfctype, comments, typ, adno, paper, adver, avm, motiv, cnt, avmversion, event) VALUES(#ORDERID#, #MYID#, '#HOST#', '#USER#', #PRIORITY#, '#TEXT#', #SAPSYSTEM#, '#APP#', '#COMMENT#', '#LOGTYP#', #ADNO#, '#PAPER#', #ADVER#, '#AVM#', #MOTIV#, #LOGCOUNT#, #VNO#, '#EVENT#')"

	Private Const defaultSQL As String = "MITASQL.INI"


	Private Structure INSPOS
		Dim posNr As Integer
		Dim statINDEX3 As Integer
		Dim psiINDEX3 As Integer
		Dim txtINDEX3 As Integer
		Dim psCount As Integer
		Dim psIndexes() As Integer
		Dim psIndex As Integer
		Dim used As Boolean
	End Structure
	Private Structure MOTIV
		Dim moNr As Integer
		Dim statINDEX7 As Integer
		Dim txtINDEX7 As Integer
		Dim psiINDEX7 As Integer
		Dim papCount As Integer
		Dim papIndexes() As Integer
		Dim papIndex As Integer
		Dim used As Boolean
	End Structure

	Private Structure EINPOSMOTIV
		Dim einNr As Integer
		Dim posNr As Integer
		Dim moNr As Integer
		Dim level As Integer
		Dim used As Boolean
	End Structure

	Private Structure EINPOS
		Dim einNr As Integer
		Dim posNr As Integer
		Dim used As Boolean
	End Structure

	Private Structure POS
		Dim posNr As Integer
		Dim statINDEX3 As Integer
		Dim psiINDEX3 As Integer
		Dim txtINDEX3 As Integer
		Dim used As Boolean
	End Structure

	Private Structure INSERTION
		'Dim pakINDEX As Integer
		Dim papIndex As Integer
		'Dim bpzINDEX As Integer
		'Dim iszINDEX As Integer
		Dim moINDEX As Integer
		Dim blzINDEX As Integer
		Dim plzINDEX As Integer
		Dim plzaINDEX As Integer
		Dim statINDEX8 As Integer
		Dim txtINDEX8 As Integer
		Dim psiINDEX8 As Integer
		'Dim errINDEX As Integer
		Dim einNr As Integer
		Dim posNr As Integer
		Dim moNr As Integer
		Dim used As Boolean
	End Structure

	Private Structure errorStructure
		Dim eAction As String
		Dim eParameter As String
	End Structure

	Private actPakINDEX As Integer = -1
	Private actPapIndex As Integer = -1
	Private actBpzINDEX As Integer = -1
	Private actIszINDEX As Integer = -1
	Private actMoINDEX As Integer = -1
	Private actBlzINDEX As Integer = -1
	Private actPsINDEX As Integer = -1
	Private actPlzINDEX As Integer = -1
	Private actPlzaINDEX As Integer = -1
	Private actStatINDEX As Integer = -1
	Private actTxtINDEX As Integer = -1
	Private actPsiINDEX As Integer = -1
	Private actErrINDEX As Integer = -1

	'PAK = 1 festes Array (0)
	Private pakCount As Integer

	'PAP = 2
	Private papCount As Integer
	Private PAP() As INSPOS

	'BPZ = 3
	Private bpzCount As Integer
	Private BPZ() As POS

	'ISZ = 4
	Private iszCount As Integer
	Private ISZ() As POS

	'MO = 5
	Private moCount As Integer
	Private moCountUsed As Integer
	Private MO() As MOTIV

	Private blzCount As Integer
	Private BLZ() As MOTIV

	'PS = 7
	Private psCount As Integer
	Private PS() As INSERTION

	'PLZ = 8
	Private plzCount As Integer
	Private PLZ() As EINPOS

	'PLZA = 9
	Private plzaCount As Integer
	Private PLZA() As EINPOS

	'STAT = 10
	Private statCount As Integer
	Private STAT() As EINPOSMOTIV

	'TXT = 11
	Private txtCount As Integer
	Private Txt() As EINPOSMOTIV

	'PSI = 12
	Private psiCount As Integer
	Private PSI() As EINPOSMOTIV

	''ERRS = 13
	'Private errCount As Integer
	'Private ERRS() As EINPOSMOTIV

	Private comboCount As Integer
	Private comboIndex As Integer
	Private combo() As String
	Private mvarPaper As String
	Private mvarCombo As String
	Private mvarInsertionCount As Integer

	Private mvarFldProduct As MITAFIELD
	Private mvarFldOrderAVM As MITAFIELD
	Private mvarFldOrderMotiv As MITAFIELD
	Private mvarFldOrderVNO As MITAFIELD
	Private mvarRefPAPAVM As MITAFIELD
	Private mvarRefPAPPos As MITAFIELD
	Private mvarFldPAPPos As MITAFIELD
	Private mvarFldPSPos As MITAFIELD
	Private mvarFldPSMo As MITAFIELD
	Private mvarFldPSEin As MITAFIELD
	Private mvarFldSTATPos As MITAFIELD
	Private mvarFldSTATMo As MITAFIELD
	Private mvarFldSTATEin As MITAFIELD
	Private mvarFldTXTPos As MITAFIELD
	Private mvarFldTXTMo As MITAFIELD
	Private mvarFldTXTEin As MITAFIELD
	Private mvarFldPSIPos As MITAFIELD
	Private mvarFldPSIMo As MITAFIELD
	Private mvarFldPSIEin As MITAFIELD
	Private mvarFldMOMo As MITAFIELD
	Private mvarFldBLZMo As MITAFIELD
	Private mvarFldBPZPos As MITAFIELD
	Private mvarFldISZPos As MITAFIELD
	Private mvarFldPLZEin As MITAFIELD
	Private mvarFldPLZPos As MITAFIELD
	Private mvarFldPLZAEin As MITAFIELD
	Private mvarFldPLZAPos As MITAFIELD
	Private mvarFldSTATLevel As MITAFIELD
	Private mvarFldTXTLevel As MITAFIELD

	Private mvarSapRecords() As MITATABLE
	Private mvarOrderBytes() As Byte
	Private mvarOrderByteCount As Integer
	Private mvarErrorBytes() As Byte
	Private mvarErrorByteCount As Integer
	Private mvarHashCode(1023) As Byte
	Private mvarSAPPakRec(0) As String
	Private mvarSAPMoRec() As String
	Private mvarSAPBlzRec() As String
	Private mvarSAPPsRec() As String
	Private mvarSAPPlzRec() As String
	Private mvarSAPPlzaRec() As String
	Private mvarSAPPsiRec() As String
	Private mvarSAPIszRec() As String
	Private mvarSAPStatRec() As String
	Private mvarSAPTxtRec() As String
	Private mvarSAPBpzRec() As String
	Private mvarSAPPapRec() As String
	'Private mvarSAPErrRec() As String
	'Private recIndex() As Integer
	Private mvarSAPRecord() As String
	Private mvarSAPRecLen() As Integer
	Private mvarSAPRecName() As String
	Private mvarSAP2DB() As MITAFIELD
	Private mvarActField As MITAFIELD
	Private mvarStartOrderTime As Double
	Private mvarStartErrorTime As Double
	Private mvarStartMotivTime As Double
	Private mvarStartReadTime As Double
	Private mvarStatTableBegin As Integer
	Private mvarStatLabelLength As Integer
	Private mvarStatContentLength As Integer
	Private mvarCustomQueries() As String
	Private mvarCustomQueryNames() As String
	Private mvarLogCount As Integer = 0
	Private mvarOldTime As Date
	Private mvarActTime As Date
	'Private aTimer As New System.Timers.Timer

	Private mvarOrderNo As String
	'Private mvarSapNo As String
	Private mvarMotivNo As String
	Private mvarPosNo As String
	Private mvarEinNo As String
	Private mvarVNO As String
	Private mvarRefAVM As String
	Private mvarRefPos As String
	Private mvarMitaSqlIni As String
	Private mvarParentPath As String
	Private mvarMyIDFlag As String
	Private createHostFlag As String
	Private mvarGarbageRemove As Boolean
	Private mvarIniFromDB As Boolean
	Private mvarAllowOlderVersion As Boolean
	Private mvarAllowSameVersion As Boolean
	Private mvarIsNew As Boolean
	Private mvarIsNewerVersion As Boolean
	Private mvarIsSameVersion As Boolean
	Private mvarIsOlderVersion As Boolean
	Private mvarNumberProcessed As Integer
	Private mvarIsRetry As Boolean = False
	Private mvarSaveAllSAP As Boolean
	Private mvarTpAdNo As String
	Private mvarTpAdVer As String
	Private mvarTpAllAd As String
	Private mvarTpNumAd As Integer
	Private mvarIniLoaded As Boolean
	Private mvarIniIsLoading As Boolean = False
	Private mvarOrderOpen As Integer
	Private mvarErrorOpen As Integer
	Private mvarOrderError As mitaErrorCodes
	Private mvarOrderInitialized As Boolean
	Private hasOrderBytes As Boolean
	Private mvarOrderID As Integer
	Private mvarMyStatus As String
	Private mvarPriority As Integer
	Private mvarText As String
	Private mvarLogTyp As String
	Private mvarComment As String
	Private mvarLogClassMap As Integer
	Private mvarLogClass As mitaSqlClass
	Private mvarNewData As String
	Private mvarNewControl As String
	Private mvarWriteBlobs As String
	Private mvarUpdateControl As String
	Private mvarRetryControl As String
	Private mvarFind As String
	Private mvarDeactivate As String
	Private mvarNewLog As String
	Private mvarGetVersion As String
	Private mvarGetID As String
	Private mvarGetAvm As String
	Private mvarGetVno As String
	Private mvarGetByVersion As String

	Private mvarActualControl As String
	Private mvarActualData As String
	Private mvarActualLength As Integer

	Private mvarSelect As String
	Private mvarSelectCount As Integer
	Private mvarSelectCursor As Integer
	Private mvarSelectID() As Integer
	'Private mvarProcessID As Integer
	'Private mvarCaption As String
	'Private mvarCommandLine As String

	Private mvarUserLogSQL() As errorStructure
	Private mvarUserLogSQLCount As Integer
	Private mvarIsSqlLog As Boolean
	Private mvaruserSqlException() As errorStructure
	Private mvaruserSqlExceptionCount As Integer
	Private mvarUserException2() As errorStructure
	Private mvarUserException2Count As Integer
	Private mvaruserException1() As errorStructure
	Private mvaruserException1Count As Integer
	Private mvarErrorProgrammer() As errorStructure
	Private mvarErrorProgrammerCount As Integer
	Private mvarErrorUser() As errorStructure
	Private mvarErrorUserCount As Integer
	Private mvarErrorSAPData() As errorStructure
	Private mvarErrorSAPDataCount As Integer
	Private mvarErrorSAPReference() As errorStructure
	Private mvarErrorSAPReferenceCount As Integer
	Private mvarErrorSAPVersion() As errorStructure
	Private mvarErrorSAPVersionCount As Integer
	Private mvarErrorSAPConnection() As errorStructure
	Private mvarErrorSAPConnectionCount As Integer
	Private mvarErrorFileSystem() As errorStructure
	Private mvarErrorFileSystemCount As Integer
	Private mvarErrorNoOpenOrder() As errorStructure
	Private mvarErrorNoOpenOrderCount As Integer
	Private mvarErrorDataBase() As errorStructure
	Private mvarErrorDataBaseCount As Integer
	Private mvarErrorDatabaseRead() As errorStructure
	Private mvarErrorDatabaseReadCount As Integer
	Private mvarErrorDatabaseConnect() As errorStructure
	Private mvarErrorDatabaseConnectCount As Integer
	Private mvarErrorNoActualSelect() As errorStructure
	Private mvarErrorNoActualSelectCount As Integer
	Private mvarErrorNoRowsSelected() As errorStructure
	Private mvarErrorNoRowsSelectedCount As Integer
	Private mvarErrorNoHost() As errorStructure
	Private mvarErrorNoHostCount As Integer
	Private mvarErrorLoop() As errorStructure
	Private mvarErrorLoopCount As Integer
	Private mvarErrorRequeryYes() As errorStructure
	Private mvarErrorRequeryYesCount As Integer
	Private mvarErrorRequeryNo() As errorStructure
	Private mvarErrorRequeryNoCount As Integer
	Private mvarErrorRequery() As errorStructure
	Private mvarErrorRequeryCount As Integer
	Private mvarErrorTrialsExceeded() As errorStructure
	Private mvarErrorTrialsExceededCount As Integer
	Private mvarErrorTrialsPossible() As errorStructure
	Private mvarErrorTrialsPossibleCount As Integer
	Private mvarErrorInvalidInput() As errorStructure
	Private mvarErrorInvalidInputCount As Integer
	Private mvarInformationEndOfList() As errorStructure
	Private mvarInformationEndOfListCount As Integer
	Private mvarUserOrderSuccess() As errorStructure
	Private mvarUserOrderSuccessCount As Integer
	Private mvarUserOrderWritten() As errorStructure
	Private mvarUserOrderWrittenCount As Integer
	Private mvarUserOrderFailure() As errorStructure
	Private mvarUserOrderFailureCount As Integer
	Private mvarUserOrderRead() As errorStructure
	Private mvarUserOrderReadCount As Integer
	Private mvarUserMotivSuccess() As errorStructure
	Private mvarUserMotivSuccessCount As Integer
	Private mvarProgramStart() As errorStructure
	Private mvarProgramStartCount As Integer
	Private mvarProgramEnd() As errorStructure
	Private mvarProgramEndCount As Integer
	Private mvarEventNames() As String
	Private mvarEventName As String
	Private mvarBreakFlag As Boolean
	Private mvarLastSqlName As String
	Private mvarLastSqlContent As String
	Private mvarLastSqlCombo As String
	Private mvarLastError As String
	Private mvarRequeryReturn As mitaEventReturnCodes
	Private mvarRequeryCount As Integer = 0

	Private mvarCustTablesArray() As custTableArray
	Private mvarCustTableCount As Integer = -1
	Private mvarOrderFrame As sapFrame
	Private mvarSapServer As pscSapServer.CSapServer

	Private mvarSapError As PscMitaError.CMitaError
	Private mvarIsTools As Boolean = False

	Private mvarOrderLogs As New ArrayList
	Private mvarNeedsLogs As Boolean

	Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
	Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer

	Public Event userEvent3(ByRef parameter As String)
	Public Event userEvent2(ByRef parameter As String)
	Public Event userEvent1(ByRef parameter As String)
	Public Event endApplication(ByRef immediatelly As Boolean)
	Public Event abortOrder()
	Public Event logContent(ByRef logText As String, ByVal logType As String)
	Public Event sapErrorForSend(ByRef sendDirect As Boolean)
	Public Event orderArrived(ByVal saveToDb As Boolean)
	Public Event sqlError(ByVal description As String, ByVal query As String, ByVal title As String)

	Private Function checkInputs() As Boolean
		checkInputs = False
		If mitaData.createID = 99999 Then
			eventProcess("Program ID not defined (namesSetID)", (mitaEventCodes.errorProgrammer), "checkInputs")
			Exit Function
		End If
		If mitaData.createUser = "" Then
			eventProcess("User not defined (namesSetUser)", (mitaEventCodes.errorProgrammer), "checkInputs")
			Exit Function
		End If
		If mitaData.mitaApplication = "" Then
			eventProcess("Application not defined (namesSetApplication)", (mitaEventCodes.errorProgrammer), "checkInputs")
			Exit Function
		End If
		checkInputs = True
	End Function

	Private Function checkOpen() As Boolean
		checkOpen = False
		If mvarOrderOpen = 0 Then
			eventProcess("No Open Order", (mitaEventCodes.errorNoOpenOrder), "checkOpen")
			Exit Function
		End If
		checkOpen = True
	End Function

	Private Function checkPath(ByRef fileName As String) As String
		Dim x() As String
		Dim name As String
		Dim i As Integer
		Dim filePath As String
		If Left$(fileName, 2) = ".\" Then
			name = Replace(fileName, ".\", Application.StartupPath & "\")
		ElseIf Mid$(fileName, 2, 1) <> ":" Then
			name = Application.StartupPath & "\" & fileName
		End If
		x = Split(name, "\")
		filePath = x(0)
		For i = 1 To UBound(x) - 1
			filePath = filePath & "\" & x(i)
			If Dir(filePath, FileAttribute.Directory) = "" Then MkDir(filePath)
		Next i
		checkPath = fileName
	End Function

	Private Sub eventDoAction(ByRef description As String, ByRef code As mitaEventCodes, ByRef eModule As String, ByRef errorHandler() As errorStructure, ByRef count As Integer)
		Dim i As Integer
		Dim decoded As String
		Dim endFlag As Boolean = False
		Dim immediateFlag As Boolean = False
		Dim abortFlag As Boolean = False
		Dim retryFlag As Boolean = False
		Dim requeryFlag As Boolean = False
		Dim normalFlag = True
		decoded = ""
		If mvarIsTools Then
			If code < mitaEventCodes.programOrderRead Then
				decoded = eventFillData("$ERRD$", description, code, eModule)
				eventPOPUP(code, decoded)
			Else
				For i = 0 To count
					Select Case errorHandler(i).eAction
						Case "LOGFILE"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventLOGFILE(decoded)
						Case "LOGDB"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventLOGDB(decoded)
							mvarComment = ""
						Case "ORDERSTART"
							eventORDERSTART()
						Case "ORDERFINISH"
							eventORDERFINISH()
						Case "ONLINELOGIN"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventLOGIN(decoded)
						Case "ONLINELOGOUT"
							eventLOGOUT()
						Case "UPDATEORDER"
							decoded = eventFillData(errorHandler(i).eParameter, Replace(description, "'", "~"), code, eModule)
							eventUPDATEORDER(decoded)
					End Select
				Next i
			End If
		Else
			For i = 0 To count
				Select Case errorHandler(i).eAction
					Case "RETRY"
						retryFlag = True
						decoded = errorHandler(i).eParameter
					Case "REQUERY"
						requeryFlag = True
						decoded = errorHandler(i).eParameter
				End Select
			Next i
			If retryFlag Then
				normalFlag = False
				Dim trials As Integer = CInt(decoded)
				If mvarNumberProcessed < trials Then
					eventProcess(description, mitaEventCodes.errorTrialsPossible, "")
					Dim sav As Boolean = mvarAllowSameVersion
					mvarAllowSameVersion = True
					mvarNumberProcessed = mvarNumberProcessed + 1
					mvarIsRetry = True
					orderWriteDB(trials)					' trials here just a dummy to receive size
					mvarIsRetry = False
					mvarNumberProcessed = mvarNumberProcessed - 1
					mvarAllowSameVersion = sav
				Else
					eventProcess(description, mitaEventCodes.errorTrialsExceeded, "")
				End If
			ElseIf requeryFlag Then
				normalFlag = False
				Dim x() As String = Split(decoded, "|")
				Dim trials As Integer = CInt(x(0))
				Dim tim As Integer = CInt(x(1))
				Dim reQuery As Boolean = False
				For i = 2 To UBound(x)
					If InStr(description, x(i), CompareMethod.Text) > 0 Then
						reQuery = True
						Exit For
					End If
				Next
				If mvarRequeryCount < trials And reQuery Then
					eventProcess(description, mitaEventCodes.errorRequeryYes, "")
					Sleep(tim)
					mvarRequeryReturn = mitaEventReturnCodes.requeryRequest
					mvarRequeryCount = mvarRequeryCount + 1
				Else
					If mvarRequeryCount >= trials Then
						eventProcess(description, mitaEventCodes.errorRequeryNo, "")
						mvarRequeryReturn = mitaEventReturnCodes.requeryExceeded
					Else
						normalFlag = True
					End If
				End If
			End If
			If normalFlag Then
				For i = 0 To count
					Select Case errorHandler(i).eAction
						Case "BREAK"
							mvarBreakFlag = True
						Case "ABORTORDER"
							abortFlag = True
							RaiseEvent abortOrder()
						Case "RAISE_END"
							If errorHandler(i).eParameter <> "" Then
								immediateFlag = CBool(errorHandler(i).eParameter)
							Else
								immediateFlag = False
							End If
							endFlag = True
						Case "MAILTO"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventMAILTO(decoded)
						Case "ALARM"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventALARM(decoded)
						Case "ABORTMOTIV"
							psCount = UBound(PS)
							papCount = UBound(PAP)
						Case "POPUP"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventPOPUP(code, decoded)
						Case "LOGFILE"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventLOGFILE(decoded)
						Case "LOGDB"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventLOGDB(decoded)
							mvarComment = ""
						Case "DUMP"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventDUMP(decoded)
						Case "SAPERR"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventSAPERR(decoded)
						Case "ORDERSTART"
							eventORDERSTART()
						Case "ORDERFINISH"
							eventORDERFINISH()
						Case "ONLINELOGIN"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventLOGIN(decoded)
						Case "ONLINELOGOUT"
							eventLOGOUT()
						Case "ORDERTOFILE"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventFILE(decoded)
						Case "RAISE_SAP"
							eventSAPERRORS(errorHandler(i).eParameter)
						Case "RAISE_EVENT1"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							RaiseEvent userEvent1(decoded)
						Case "RAISE_EVENT2"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							RaiseEvent userEvent2(decoded)
						Case "RAISE_EVENT3"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							RaiseEvent userEvent3(decoded)
						Case "RAISE_LOG"
							decoded = eventFillData(errorHandler(i).eParameter, description, code, eModule)
							eventRAISELOG(decoded)
						Case "UPDATEORDER"
							decoded = eventFillData(errorHandler(i).eParameter, Replace(description, "'", "~"), code, eModule)
							eventUPDATEORDER(decoded)
						Case "RETRY"
						Case "REQUERY"
						Case Else
							eventPOPUP(mitaEventCodes.errorProgrammer, "Unhandled Action:" & vbCrLf & errorHandler(i).eAction)
					End Select
				Next i
			End If
			If code < mitaEventCodes.programOrderRead And description <> "" Then
				mvarLastError = Replace(description, "'", "~")
			End If
			If abortFlag Then
				If Not IsNothing(MO) Then moCount = UBound(MO)
				If Not IsNothing(PS) Then psCount = UBound(PS)
				If Not IsNothing(PAP) Then papCount = UBound(PAP)
			End If
		End If
		If endFlag Then
			RaiseEvent endApplication(immediateFlag)
		End If
	End Sub

	Private Sub eventALARM(ByRef action As String)
		Dim Where As String = " WHERE sapsystemid = " & mitaSystem.sapSystemId
		Where = Where & " AND loginid = " & mitaData.createID
		Where = Where & " AND loginhost = '" & mitaData.createHost & "'"
		Where = Where & " AND loginapp = '" & mitaData.mitaApplication & "'"
		Dim query As String = "UPDATE " & mitaSystem.tableOnline
		query = query & " SET alarm = '" & action & "'"
		sqlExec(query & Where)
	End Sub

	Private Sub eventUPDATEORDER(ByRef action As String)
		Dim query As New System.Text.StringBuilder(1000)
		Dim x() As String
		Dim i As Integer
		Dim result As Boolean
		query.Append("UPDATE " & mitaSystem.tableOrderControl & " SET ")
		x = Split(action, "|")
		For i = 0 To UBound(x)
			If i > 0 Then query.Append(", ")
			query.Append(x(i))
		Next
		query.Append(" WHERE pscid = " & mvarOrderID)
		result = sqlExecTrans(query.ToString)
	End Sub
	Private Function eventFillData(ByRef src As String, ByRef description As String, ByRef code As mitaEventCodes, ByRef eModule As String) As String
		Dim codeText As String
		Dim result As New System.Text.StringBuilder(src, 1000)
		Dim dat As String
		Dim tim As String
		Dim motivTime As String
		Dim readTime As String
		Dim orderTime As String
		Dim errorTime As String
		Dim resLen As Integer
		Dim errLen As Integer
		dat = FormatDateTime(Now, DateFormat.ShortDate)
		tim = FormatDateTime(Now, DateFormat.LongTime)
		If InStr(src, "$ERRT$") <> 0 Then
			codeText = mvarEventNames(code)
			result.Replace("$ERRT$", codeText)
		End If
		result.Replace("$ERRC$", CStr(code))
		If InStr(src, "$ERRM$") <> 0 Then
			If eModule <> "" Then
				result.Replace("$ERRM$", "(" & eModule & ")")
			Else
				result.Replace("$ERRM$", "")
			End If
		End If
		result.Replace("$VNO$", CStr(mvarVNO))
		result.Replace("$AVM$", mvarOrderNo)
		result.Replace("$MO$", CStr(mvarMotivNo))
		result.Replace("$PSC$", CStr(mvarOrderID))
		result.Replace("$HOST$", mitaData.createHost)
		result.Replace("$ID$", CStr(mitaData.createID))
		result.Replace("$DATE$", dat)
		result.Replace("$TIME$", tim)
		If InStr(src, "$MTIM$") <> 0 Then
			motivTime = Format(VB.Timer() - mvarStartMotivTime, "000.000")
			result.Replace("$MTIM$", motivTime)
		End If
		If InStr(src, "$RTIM$") <> 0 Then
			readTime = Format(VB.Timer() - mvarStartReadTime, "000.000")
			result.Replace("$RTIM$", readTime)
		End If
		If InStr(src, "$OTIM$") <> 0 Then
			orderTime = Format(VB.Timer() - mvarStartOrderTime, "000.000")
			result.Replace("$OTIM$", orderTime)
		End If
		If InStr(src, "$ETIM$") <> 0 Then
			errorTime = Format(VB.Timer() - mvarStartErrorTime, "000.000")
			result.Replace("$ETIM$", errorTime)
		End If
		result.Replace("$AD$", mvarTpAdNo)
		result.Replace("$ADVER$", mvarTpAdVer)
		result.Replace("$PAPER$", mvarPaper)
		If mvarCombo <> "" Then
			result.Replace("$COMBO$", " (" & mvarCombo & ")")
		Else
			result.Replace("$COMBO$", "")
		End If
		If moCount > 0 Then
			moCountUsed = moCountUsed + 1
			moCountUsed = moCountUsed - 1
		End If
		result.Replace("$GARB$", CStr(moCount - moCountUsed + 1))
		result.Replace("$INS$", CStr((psCount + 1) * (comboCount + 1)))
		result.Replace("$ALLAD$", mvarTpAllAd)
		result.Replace("$NUMAD$", CStr(mvarTpNumAd))
		result.Replace("$$", vbCrLf)
		result.Replace("$SIZE$", CStr(mvarOrderByteCount))
		Dim tmp As String = result.ToString
		If Not IsNothing(description) Then
			resLen = tmp.Length
			errLen = description.Length
			If resLen + errLen <= 253 Then
				tmp = tmp.Replace("$ERRD$", description)
			Else
				tmp = tmp.Replace("$ERRD$", Left$(description, 253 - resLen))
			End If
		End If
		If Not IsNothing(mvarLastError) Then
			resLen = tmp.Length
			errLen = mvarLastError.Length
			If resLen + errLen <= 253 Then
				tmp = tmp.Replace("$ERROR$", mvarLastError)
			Else
				tmp = tmp.Replace("$ERROR$", Left$(mvarLastError, 253 - resLen))
			End If
		End If
		Return tmp
	End Function

	Private Function eventGetCodeText(ByRef code As mitaEventCodes) As String
		Dim codeText As String
		Select Case code
			Case mitaEventCodes.errorDatabaseConnect : codeText = "errorUserIni"
			Case mitaEventCodes.errorUserIni : codeText = "errorDatabaseConnect"
			Case mitaEventCodes.errorDatabaseSequence : codeText = "errorDatabaseSequence"
			Case mitaEventCodes.errorFileSystem : codeText = "errorFileSystem"
			Case mitaEventCodes.errorInvalidInput : codeText = "errorInvalidInput"
			Case mitaEventCodes.errorNoActualSelect : codeText = "errorNoActualSelect"
			Case mitaEventCodes.errorNoHost : codeText = "errorNoHost"
			Case mitaEventCodes.errorLoop : codeText = "errorLoop"
			Case mitaEventCodes.errorRequery : codeText = "errorRequery"
			Case mitaEventCodes.errorRequeryYes : codeText = "errorRequeryYes"
			Case mitaEventCodes.errorRequeryNo : codeText = "errorRequeryNo"
			Case mitaEventCodes.errorTrialsExceeded : codeText = "errorTrialsExceeded"
			Case mitaEventCodes.errorTrialsPossible : codeText = "errorTrialsPossible"
			Case mitaEventCodes.errorNoOpenOrder : codeText = "errorNoOpenOrder"
			Case mitaEventCodes.errorDataBase : codeText = "errorDataBase"
			Case mitaEventCodes.errorNoRowsSelected : codeText = "errorNoRowsSelected"
			Case mitaEventCodes.errorProgrammer : codeText = "errorProgrammer"
			Case mitaEventCodes.errorSAPConnection : codeText = "errorSAPConnection"
			Case mitaEventCodes.errorSAPReference : codeText = "errorSAPReference"
			Case mitaEventCodes.errorSAPData : codeText = "errorSAPData"
			Case mitaEventCodes.errorSAPVersion : codeText = "errorSAPVersion"
			Case mitaEventCodes.errorUserIni : codeText = "errorUser"
			Case mitaEventCodes.userLogSQL : codeText = "userLogSQL"
			Case mitaEventCodes.userSqlException : codeText = "userSqlException"
			Case mitaEventCodes.userException2 : codeText = "userException2"
			Case mitaEventCodes.userException1 : codeText = "userException1"
			Case mitaEventCodes.programOrderRead : codeText = "programOrderRead"
			Case mitaEventCodes.userMotivSuccess : codeText = "userMotivSuccess"
			Case mitaEventCodes.userOrderSuccess : codeText = "userOrderSuccess"
			Case mitaEventCodes.userOrderFailure : codeText = "userOrderFailure"
			Case mitaEventCodes.programOrderWritten : codeText = "programOrderWritten"
			Case mitaEventCodes.programStart : codeText = "programStart"
			Case mitaEventCodes.programEnd : codeText = "programEnd"
		End Select
		eventGetCodeText = codeText
	End Function

	Private Function eventGetCodeValue(ByVal codeText As String) As mitaEventCodes
		Dim codeVal As mitaEventCodes
		Select Case codeText
			Case "errorDatabaseConnect" : codeVal = mitaEventCodes.errorDatabaseConnect
			Case "errorUserIni" : codeVal = mitaEventCodes.errorUserIni
			Case "errorDatabaseSequence" : codeVal = mitaEventCodes.errorDatabaseSequence
			Case "errorFileSystem" : codeVal = mitaEventCodes.errorFileSystem
			Case "errorInvalidInput" : codeVal = mitaEventCodes.errorInvalidInput
			Case "errorNoActualSelect" : codeVal = mitaEventCodes.errorNoActualSelect
			Case "errorNoHost" : codeVal = mitaEventCodes.errorNoHost
			Case "errorLoop" : codeVal = mitaEventCodes.errorLoop
			Case "errorRequery" : codeVal = mitaEventCodes.errorRequery
			Case "errorRequeryYes" : codeVal = mitaEventCodes.errorRequeryYes
			Case "errorRequeryNo" : codeVal = mitaEventCodes.errorRequeryNo
			Case "errorTrialsExceeded" : codeVal = mitaEventCodes.errorTrialsExceeded
			Case "errorTrialsPossible" : codeVal = mitaEventCodes.errorTrialsPossible
			Case "errorNoOpenOrder" : codeVal = mitaEventCodes.errorNoOpenOrder
			Case "errorDataBase" : codeVal = mitaEventCodes.errorDataBase
			Case "errorNoRowsSelected" : codeVal = mitaEventCodes.errorNoRowsSelected
			Case "errorProgrammer" : codeVal = mitaEventCodes.errorProgrammer
			Case "errorSAPConnection" : codeVal = mitaEventCodes.errorSAPConnection
			Case "errorSAPReference" : codeVal = mitaEventCodes.errorSAPReference
			Case "errorSAPData" : codeVal = mitaEventCodes.errorSAPData
			Case "errorSAPVersion" : codeVal = mitaEventCodes.errorSAPVersion
			Case "errorUser" : codeVal = mitaEventCodes.errorUserIni
			Case "userLogSQL" : codeVal = mitaEventCodes.userLogSQL
			Case "userSqlException" : codeVal = mitaEventCodes.userSqlException
			Case "userException2" : codeVal = mitaEventCodes.userException2
			Case "userException1" : codeVal = mitaEventCodes.userException1
			Case "programOrderRead" : codeVal = mitaEventCodes.programOrderRead
			Case "userMotivSuccess" : codeVal = mitaEventCodes.userMotivSuccess
			Case "userOrderSuccess" : codeVal = mitaEventCodes.userOrderSuccess
			Case "userOrderFailure" : codeVal = mitaEventCodes.userOrderFailure
			Case "programOrderWritten" : codeVal = mitaEventCodes.programOrderWritten
			Case "programStart" : codeVal = mitaEventCodes.programStart
			Case "programEnd" : codeVal = mitaEventCodes.programEnd
		End Select
		Return codeVal
	End Function


	Private Sub eventDUMP(ByRef filename As String)
		Dim i As Integer
		Dim f1 As Integer
		Dim a As String
		Dim itm As SQLLOG
		If filename <> "" Then
			filename = checkPath(filename)
			f1 = FreeFile()
			FileOpen(f1, filename, OpenMode.Output)
			itm = mvarOrderLogs(0)
			For i = 0 To mvarOrderLogs.Count - 1
				itm = CType(mvarOrderLogs(i), SQLLOG)
				PrintLine(f1, itm.sName)
				If itm.sError <> "" Then PrintLine(f1, itm.sError)
				PrintLine(f1, itm.sTemplate)
				PrintLine(f1, itm.sResult)
				PrintLine(f1, "")
			Next i
			FileClose(f1)
		End If
	End Sub
	Private Sub eventLOGFILE(ByRef parameter As String)
		Dim x() As String
		Dim i As Integer
		Dim max As Integer
		Dim f1 As Integer
		Dim f2 As Integer
		Dim fileName As String
		Dim logText As String
		Dim a As String
		x = Split(parameter, "|")
		fileName = x(0)
		logText = x(1)
		If fileName <> "" Then
			fileName = checkPath(fileName)
			f1 = FreeFile()
			FileOpen(f1, fileName, OpenMode.Append)
			PrintLine(f1, logText)
			FileClose(f1)
			If UBound(x) = 2 Then
				max = 1000 * CInt(x(2))
				If FileLen(fileName) > max Then
					f1 = FreeFile()
					FileOpen(f1, fileName, OpenMode.Input)
					f2 = FreeFile()
					FileOpen(f1, Application.StartupPath & "\log.tmp", OpenMode.Output)
					For i = 0 To 100
						If EOF(f1) Then Exit For
						a = LineInput(f1)
					Next i
					Do
						If EOF(f1) Then Exit Do
						a = LineInput(f1)
						PrintLine(f2, a)
					Loop
					FileClose(f1)
					FileClose(f2)
					Kill(fileName)
					Rename(Application.StartupPath & "\log.tmp", fileName)
				End If
			End If
		End If
	End Sub

	Private Sub eventLOGDB(ByRef parameter As String)
		Dim x() As String
		x = Split(parameter, "|")
		mvarComment = ""
		mvarPriority = 0
		mvarText = Replace(x(0), "'", "~")
		If x(1) <> "S" Then
			mvarLogTyp = x(1)
		Else
			mvarLogTyp = mvarLogTyp & x(1)
		End If
		If UBound(x) > 1 Then mvarPriority = CShort(x(2))
		If UBound(x) > 2 Then mvarComment = x(3)
		orderWriteDBLog()
	End Sub

	Private Sub eventRAISELOG(ByRef parameter As String)
		Dim x() As String
		x = Split(parameter, "|")
		mvarText = x(0)
		If UBound(x) = 0 Then
			mvarLogTyp = "L"
		Else
			mvarLogTyp = x(1)
		End If
		RaiseEvent logContent(mvarText, mvarLogTyp)
	End Sub

	Private Sub eventLOGIN(ByRef parameter As String)
		Dim query As String
		Dim Where As String
		Dim q As String
		Dim number As Integer
		Dim result As Boolean
		Dim mail As String
		Dim x() As String
		Dim lastOrder As String
		Dim actOrder As String
		Dim loopReported As Boolean = False
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader

		query = "SELECT count(*) FROM " & mitaSystem.tableOnline
		Where = " WHERE sapsystemid = " & mitaSystem.sapSystemId
		Where = Where & " AND loginid = " & mitaData.createID
		Where = Where & " AND loginhost = '" & mitaData.createHost & "'"
		Where = Where & " AND loginapp = '" & mitaData.mitaApplication & "'"
		result = mitaConnect.queryNumber(query & Where, number)
		If result Then
			If number = 0 Then
				query = "INSERT INTO " & mitaSystem.tableOnline & " (sapsystemid, loginid, loginhost, loginapp, processid, caption, command, hostindex)"
				query = query & " VALUES("
				query = query & CStr(mitaSystem.sapSystemId)
				query = query & ", " & mitaData.createID
				query = query & ", '" & mitaData.createHost & "'"
				query = query & ", '" & mitaData.mitaApplication & "'"
				query = query & ", " & mitaData.processID
				query = query & ", '" & mitaData.caption & "'"
				query = query & ", '" & mitaData.commandLine & "'"
				query = query & ", " & mitaData.newID
				query = query & ")"
				result = sqlExec(query)
			Else
				query = "SELECT logouttime, orderno, lastorder FROM " & mitaSystem.tableOnline
				q = query & Where
				idbc.CommandText = query
				mitaConnect.odbc_connection.Open()
				reader = idbc.ExecuteReader()
				If reader.Read Then
					If IsDBNull(reader.GetValue(0)) Then
						If Not IsDBNull(reader.GetValue(2)) Then
							lastOrder = reader.GetString(2)
						Else
							lastOrder = ""
						End If
						If Not IsDBNull(reader.GetValue(1)) Then
							actOrder = reader.GetString(1)
						Else
							actOrder = ""
						End If
						reader.Close()
						idbc.Dispose()
						mitaConnect.odbc_connection.Close()
						If actOrder = lastOrder And actOrder <> "" Then
							mvarOrderID = CInt(actOrder)
							mitaConnect.queryString(sqlTrans(mvarGetAvm), mvarOrderNo)
							Dim tmp As Integer
							mitaConnect.queryNumber(sqlTrans(mvarGetVno), tmp)
							mvarVNO = CStr(tmp)
							eventRaise(actOrder, mitaEventCodes.errorLoop)
							loopReported = True
						End If
					Else
						reader.Close()
						idbc.Dispose()
						mitaConnect.odbc_connection.Close()
					End If
				Else
					reader.Close()
					idbc.Dispose()
					mitaConnect.odbc_connection.Close()
				End If
				query = "UPDATE " & mitaSystem.tableOnline
				If Not loopReported Then
					query = query & " SET orderno = lastorder"
				Else
					query = query & " SET orderno = NULL"
					query = query & ", lastorder = NULL"
				End If
				query = query & ", logintime = sysdate"
				query = query & ", logouttime = NULL"
				query = query & ", alive = sysdate"
				query = query & ", alarm = NULL"
				query = query & ", processid = " & mitaData.processID
				query = query & ", caption = '" & mitaData.caption & "'"
				query = query & ", command = '" & mitaData.commandLine & "'"
				result = sqlExec(query & Where)
			End If
			mail = parameter
		End If
	End Sub
	Private Sub eventORDERSTART()
		Dim query As New System.Text.StringBuilder(1000)
		Dim Where As String
		Dim result As Boolean
		query.Append("UPDATE " & mitaSystem.tableOnline)
		query.Append(" SET orderno = '" & mvarOrderID & "', alive = sysdate")
		query.Append(" WHERE sapsystemid = " & mitaSystem.sapSystemId)
		query.Append(" AND loginid = " & mitaData.createID)
		query.Append(" AND loginhost = '" & mitaData.createHost & "'")
		query.Append(" AND loginapp = '" & mitaData.mitaApplication & "'")
		query.Append(";")
		result = sqlExec(query.ToString)
	End Sub
	Private Sub eventORDERFINISH()
		Dim query As New System.Text.StringBuilder(1000)
		Dim result As Boolean
		query.Append("UPDATE " & mitaSystem.tableOnline)
		query.Append(" SET lastorder = orderno, alive = sysdate")
		query.Append(", orderno = NULL")
		query.Append(" WHERE sapsystemid = " & mitaSystem.sapSystemId)
		query.Append(" AND loginid = " & mitaData.createID)
		query.Append(" AND loginhost = '" & mitaData.createHost & "'")
		query.Append(" AND loginapp = '" & mitaData.mitaApplication & "'")
		result = sqlExec(query.ToString)
	End Sub

	Private Sub eventLOGOUT()
		Dim query As String
		Dim Where As String
		Where = " WHERE sapsystemid = " & mitaSystem.sapSystemId
		Where = Where & " AND loginid = " & mitaData.createID
		Where = Where & " AND loginhost = '" & mitaData.createHost & "'"
		Where = Where & " AND loginapp = '" & mitaData.mitaApplication & "'"
		query = "UPDATE " & mitaSystem.tableOnline
		query = query & " SET logouttime = sysdate, alive = NULL, processid = NULL, caption = NULL, command = NULL, orderno = NULL, lastorder = NULL"
		sqlExec(query & Where)
	End Sub

	Private Sub eventFILE(ByRef parameter As String)
		Dim fileName As String
		Dim result As Boolean
		Dim max As Integer
		fileName = checkPath(parameter)
		If Dir$(fileName) <> "" Then Kill(fileName)
		result = orderWriteFile(fileName, max)
	End Sub
	Private Sub eventMAILTO(ByRef parameter As String)
		Dim x() As String
		Dim a As Integer
		x = Split(parameter, "|")
		'If UBound(x) = 2 Then
		'frmMitaMsgInst.sendMail("pscMitaOrder@" & mitaData.createHost, x(0), x(1), x(2))
		Dim bAns As Boolean = True
		Dim sParams As String
		sParams = "mailto:" & x(0)
		sParams = sParams & "?subject=" & x(1) & "?pscMitaOrder@" & mitaData.createHost & vbCrLf & x(2)
		Try
			System.Diagnostics.Process.Start(sParams)
		Catch
			bAns = False
		End Try
		'ElseIf UBound(x) = 3 Then
		'	If UCase(Trim(x(3))) = "TRUE" Or CInt(x(3)) = 1 Then
		'		orderWriteFile(mvarParentPath & "mail", a)
		'		frmMitaMsgInst.sendMail("pscMitaOrder@" & mitaData.createHost, x(0), x(1), x(2), mvarParentPath & "mail")
		'		Kill(mvarParentPath & "mail")
		'	Else
		'		frmMitaMsgInst.sendMail("pscMitaOrder@" & mitaData.createHost, x(0), x(1), x(2))
		'	End If
		'End If
	End Sub

	Private Sub eventPOPUP(ByRef code As mitaEventCodes, ByRef parameter As String)
		Dim x() As String
		If mvarEventNames Is Nothing Then
			If InStr(parameter, "|") = 0 Then
				mitaMessage.waitTime = 0
				mitaMessage.popupMessage(parameter, "")
			Else
				x = Split(parameter, "|")
				mitaMessage.waitTime = CShort(x(1))
				mitaMessage.popupMessage(x(0), "")
			End If
		Else
			If InStr(parameter, "|") = 0 Then
				mitaMessage.waitTime = 0
				mitaMessage.popupMessage(parameter, mvarEventNames(code))
			Else
				x = Split(parameter, "|")
				mitaMessage.waitTime = CShort(x(1))
				mitaMessage.popupMessage(x(0), mvarEventNames(code))
			End If
		End If
	End Sub

	Private Sub eventProcess(ByRef description As String, ByRef code As mitaEventCodes, ByRef eModule As String)
		Dim x() As String
		mvarEventName = mvarEventNames(code)
		mitaConnect.connectPush()
		Select Case code
			Case mitaEventCodes.errorProgrammer : eventDoAction(description, code, eModule, mvarErrorProgrammer, mvarErrorProgrammerCount)
			Case mitaEventCodes.errorUserIni : eventDoAction(description, code, eModule, mvarErrorUser, mvarErrorUserCount)
			Case mitaEventCodes.errorSAPConnection : eventDoAction(description, code, eModule, mvarErrorSAPConnection, mvarErrorSAPConnectionCount)
			Case mitaEventCodes.errorSAPReference : eventDoAction(description, code, eModule, mvarErrorSAPReference, mvarErrorSAPReferenceCount)
			Case mitaEventCodes.errorSAPData : eventDoAction(description, code, eModule, mvarErrorSAPData, mvarErrorSAPDataCount)
			Case mitaEventCodes.errorSAPVersion : eventDoAction(description, code, eModule, mvarErrorSAPVersion, mvarErrorSAPVersionCount)
			Case mitaEventCodes.errorFileSystem : eventDoAction(description, code, eModule, mvarErrorFileSystem, mvarErrorFileSystemCount)
			Case mitaEventCodes.errorNoOpenOrder : eventDoAction(description, code, eModule, mvarErrorNoOpenOrder, mvarErrorNoOpenOrderCount)
			Case mitaEventCodes.errorDataBase : eventDoAction(description, code, eModule, mvarErrorDataBase, mvarErrorDataBaseCount)
			Case mitaEventCodes.errorDatabaseSequence : eventDoAction(description, code, eModule, mvarErrorDatabaseRead, mvarErrorDatabaseReadCount)
			Case mitaEventCodes.errorDatabaseConnect : eventDoAction(description, code, eModule, mvarErrorDatabaseConnect, mvarErrorDatabaseConnectCount)
			Case mitaEventCodes.errorNoActualSelect : eventDoAction(description, code, eModule, mvarErrorNoActualSelect, mvarErrorNoActualSelectCount)
			Case mitaEventCodes.errorNoRowsSelected : eventDoAction(description, code, eModule, mvarErrorNoRowsSelected, mvarErrorNoRowsSelectedCount)
			Case mitaEventCodes.errorInvalidInput : eventDoAction(description, code, eModule, mvarErrorInvalidInput, mvarErrorInvalidInputCount)
			Case mitaEventCodes.errorNoHost : eventDoAction(description, code, eModule, mvarErrorNoHost, mvarErrorNoHostCount)
			Case mitaEventCodes.errorLoop : eventDoAction(description, code, eModule, mvarErrorLoop, mvarErrorLoopCount)
			Case mitaEventCodes.errorRequery : eventDoAction(description, code, eModule, mvarErrorRequery, mvarErrorRequeryCount)
			Case mitaEventCodes.errorRequeryYes : eventDoAction(description, code, eModule, mvarErrorRequeryYes, mvarErrorRequeryYesCount)
			Case mitaEventCodes.errorRequeryNo : eventDoAction(description, code, eModule, mvarErrorRequeryNo, mvarErrorRequeryNoCount)
			Case mitaEventCodes.errorTrialsExceeded : eventDoAction(description, code, eModule, mvarErrorTrialsExceeded, mvarErrorTrialsExceededCount)
			Case mitaEventCodes.errorTrialsPossible : eventDoAction(description, code, eModule, mvarErrorTrialsPossible, mvarErrorTrialsPossibleCount)
			Case mitaEventCodes.userSqlException
				mvarComment = mvarLastSqlName & ": " & mvarLastSqlContent
				eventDoAction(description, code, eModule, mvaruserSqlException, mvaruserSqlExceptionCount)
				mvarLastSqlCombo = ""
				mvarLastSqlName = ""
				mvarLastSqlContent = ""
			Case mitaEventCodes.userException2 : eventDoAction(description, code, eModule, mvarUserException2, mvarUserException2Count)
			Case mitaEventCodes.userException1 : eventDoAction(description, code, eModule, mvaruserException1, mvaruserException1Count)
			Case mitaEventCodes.programOrderRead : eventDoAction(description, code, eModule, mvarUserOrderRead, mvarUserOrderReadCount)
			Case mitaEventCodes.programOrderWritten : eventDoAction(description, code, eModule, mvarUserOrderWritten, mvarUserOrderWrittenCount)
			Case mitaEventCodes.userMotivSuccess : eventDoAction(description, code, eModule, mvarUserMotivSuccess, mvarUserMotivSuccessCount)
			Case mitaEventCodes.userOrderSuccess : eventDoAction(description, code, eModule, mvarUserOrderSuccess, mvarUserOrderSuccessCount)
			Case mitaEventCodes.userOrderFailure : eventDoAction(description, code, eModule, mvarUserOrderFailure, mvarUserOrderFailureCount)
			Case mitaEventCodes.userLogSQL
				Dim savPaper As String = mvarPaper
				x = Split(description, "@")
				If UBound(x) <> 3 Then
					eventRaise("SQL log must have 4 parameters, separated by '@': " & vbCrLf & description, mitaEventCodes.errorProgrammer, eModule)
					Exit Sub
				End If
				mvarLogClass = CType(x(0), mitaSqlClass)
				If (mvarLogClassMap And mvarLogClass) > 0 Or (mvarLogClass = mitaSqlClass.classError) Then
					Select Case mvarLogClass
						Case mitaSqlClass.classAd
							mvarLogTyp = "A"
						Case mitaSqlClass.classInternal
							mvarLogTyp = "I"
						Case mitaSqlClass.classOrder
							mvarLogTyp = "O"
						Case mitaSqlClass.classError
							mvarLogTyp = "E"
					End Select
					mvarPaper = x(3)
					eventDoAction(x(1) & ": " & x(2), code, eModule, mvarUserLogSQL, mvarUserLogSQLCount)
					mvarPaper = savPaper
				End If
				mvarLastSqlCombo = mvarPaper
				mvarLastSqlName = x(1)
				mvarLastSqlContent = Replace(x(2), "'", "~")
			Case mitaEventCodes.programStart : eventDoAction(description, code, eModule, mvarProgramStart, mvarProgramStartCount)
			Case mitaEventCodes.programEnd : eventDoAction(description, code, eModule, mvarProgramEnd, mvarProgramEndCount)
			Case mitaEventCodes.userMessage : eventRAISELOG(eventFillData(description, "", mitaEventCodes.userMessage, ""))
		End Select
		mitaConnect.connectPop()
	End Sub

	Private Sub createVersionInfo()
		Dim result As Boolean
		mitaConnect.connectPush()
		Dim ok As Boolean = mitaConnect.queryExist(sqlTrans(mvarGetVersion), result)
		mitaConnect.connectPop()
		mvarIsNew = Not result
		mvarIsNewerVersion = False
		mvarIsSameVersion = False
		mvarIsOlderVersion = False
		If Not mvarIsNew Then
			Dim actver As Integer
			mitaConnect.connectPush()
			ok = mitaConnect.queryNumber(sqlTrans(mvarGetVersion), actver)
			mitaConnect.connectPop()
			If Not ok Then Exit Sub
			mvarIsNewerVersion = CInt(mvarVNO) > actver
			mvarIsSameVersion = CInt(mvarVNO) = actver
			mvarIsOlderVersion = CInt(mvarVNO) < actver
		End If
	End Sub

	Private Function checkVersion(ByRef result As Boolean) As Boolean
		checkVersion = mitaConnect.queryExist(sqlTrans(mvarGetVersion), result)
	End Function

	Public Function eventRaise(ByRef errorDescription As String, ByRef errorCode As mitaEventCodes, Optional ByVal base As String = "") As mitaEventReturnCodes
		Dim returnCode As mitaEventReturnCodes = mitaEventReturnCodes.returnOk
		'If Not mvarIsDBConnected Then connectionOpen (mvarDBConnectString)
		If mvarIniIsLoading Then Exit Function
		If Not mvarIniLoaded Then iniRead()
		If mvarIniLoaded Then
			eventProcess(errorDescription, errorCode, base)
		End If
		If mvarBreakFlag Then
			returnCode = returnCode Or mitaEventReturnCodes.debugBreak
			mvarBreakFlag = False
		End If
		If mvarRequeryReturn <> mitaEventReturnCodes.returnOk Then
			returnCode = returnCode Or mvarRequeryReturn
		End If
		eventRaise = returnCode
	End Function

	Public Function eventLog(ByRef logStruct As SQLLOG) As Boolean
		Dim savAdNo As String
		Dim savVerNo As String
		If Not mvarIniLoaded Then iniRead()
		If mvarIniLoaded Then
			If mvarNeedsLogs Then mvarOrderLogs.Add(logStruct)
			savAdNo = mvarTpAdNo
			savVerNo = mvarTpAdVer
			mvarTpAdNo = logStruct.sNumberAd
			mvarTpAdVer = logStruct.sVersionAd
			eventRaise(CStr(logStruct.sClass) & "@" & logStruct.sName & "@" & logStruct.sResult & "@" & logStruct.sPaper, mitaEventCodes.userLogSQL)
			mvarTpAdNo = savAdNo
			mvarTpAdVer = savVerNo
			eventLog = True
		Else
			eventLog = False
		End If
	End Function
	Public Function infoConvertAdvName(ByRef fileName As String) As String
		infoConvertAdvName = rad36Decode(fileName)
	End Function

	Public Function infoGetcombis(ByRef product As String) As String()
		Dim x() As String
		Dim i As Integer
		For i = 0 To mitaData.combiCount
			If mitaData.combis(i).cName = product Then
				x = Split(mitaData.combis(i).cEntry, ",")
				Exit For
			End If
		Next i
		If i > mitaData.combiCount Then
			ReDim x(0)
			x(0) = product
		End If
		Return x
	End Function

	Public Function infoGetCodeText(ByRef code As mitaEventCodes) As String
		infoGetCodeText = mvarEventNames(code)
	End Function

	'Public Function infoGetStructFileName() As String
	'	infoGetStructFileName = mvarMitaStructIni
	'End Function

	Public Function infoGetRecords(ByRef Index As Integer) As String()
		Dim s() As String
		Dim C As Integer
		Dim i As Integer
		C = -1
		ReDim s(0)
		For i = 0 To mvarSapRecords(Index).mtCount - 1
			If mvarSapRecords(Index).mtRecords(i).sbUsed Or Not mvarGarbageRemove Then
				C = C + 1
				ReDim Preserve s(C)
				s(C) = mvarSapRecords(Index).mtRecords(i).sbContent
			End If
		Next i
		Return s
	End Function


	Public Function infoGetRecordCount(ByRef Index As Integer) As Integer
		Dim C As Integer
		Dim i As Integer
		C = -1
		For i = 0 To mvarSapRecords(Index).mtCount - 1
			If mvarSapRecords(Index).mtRecords(i).sbUsed Or Not mvarGarbageRemove Then
				C = C + 1
			End If
		Next i
		infoGetRecordCount = C + 1
	End Function

	Public Function infoGetLength(ByRef Index As Integer) As Integer
		infoGetLength = mvarSapRecords(Index).mtLength
	End Function

	Public Function infoSetAdNo(ByRef thirdPartyAdNumber As Integer) As Boolean
		If thirdPartyAdNumber > 0 Then
			mvarTpAdNo = CStr(thirdPartyAdNumber)
			If mvarTpAllAd <> "" Then mvarTpAllAd = mvarTpAllAd & ", "
			mvarTpNumAd = mvarTpNumAd + 1
			mvarTpAllAd = mvarTpAllAd & mvarTpAdNo
		Else
			mvarTpAdNo = ""
		End If
		infoSetAdNo = True
	End Function
	Public Function infoSetAdVer(ByRef thirdPartyAdVersion As Integer) As Boolean
		If thirdPartyAdVersion > 0 Then
			mvarTpAdVer = CStr(thirdPartyAdVersion)
		Else
			mvarTpAdVer = ""
		End If
		infoSetAdVer = True
	End Function

	Private Sub iniReadIntern()
		If mvarUpdateControl = "" Then
			mvarUpdateControl = cUpdateControl
			mvarRetryControl = cRetryControl
			mvarFind = cFind
			mvarGetVersion = cGetVersion
			mvarGetID = cGetID
			mvarGetAvm = cGetAvm
			mvarGetVno = cGetVno
			mvarGetByVersion = cGetByVersion
			mvarNewData = cNewData
			mvarNewControl = cNewControl
			mvarWriteBlobs = cWriteBlobs
			mvarDeactivate = cDeactivate
			mvarNewLog = cNewLog
		End If
		iniReadErrorHandling("userLogSQL", mvarUserLogSQL, mvarUserLogSQLCount)
		iniReadErrorHandling("userSqlException", mvaruserSqlException, mvaruserSqlExceptionCount)
		iniReadErrorHandling("userException2", mvarUserException2, mvarUserException2Count)
		iniReadErrorHandling("userException1", mvaruserException1, mvaruserException1Count)
		iniReadErrorHandling("errorProgrammer", mvarErrorProgrammer, mvarErrorProgrammerCount)
		iniReadErrorHandling("errorUserIni", mvarErrorUser, mvarErrorUserCount)
		iniReadErrorHandling("errorSAPData", mvarErrorSAPData, mvarErrorSAPDataCount)
		iniReadErrorHandling("errorSAPReference", mvarErrorSAPReference, mvarErrorSAPReferenceCount)
		iniReadErrorHandling("errorSAPVersion", mvarErrorSAPVersion, mvarErrorSAPVersionCount)
		iniReadErrorHandling("errorSAPConnection", mvarErrorSAPConnection, mvarErrorSAPConnectionCount)
		iniReadErrorHandling("errorFileSystem", mvarErrorFileSystem, mvarErrorFileSystemCount)
		iniReadErrorHandling("errorNoOpenOrder", mvarErrorNoOpenOrder, mvarErrorNoOpenOrderCount)
		iniReadErrorHandling("errorNoOpenOrder", mvarErrorDataBase, mvarErrorDataBaseCount)
		iniReadErrorHandling("errorDatabaseSequence", mvarErrorDatabaseRead, mvarErrorDatabaseReadCount)
		iniReadErrorHandling("errorDatabaseConnect", mvarErrorDatabaseConnect, mvarErrorDatabaseConnectCount)
		iniReadErrorHandling("errorNoActualSelect", mvarErrorNoActualSelect, mvarErrorNoActualSelectCount)
		iniReadErrorHandling("errorNoRowsSelected", mvarErrorNoRowsSelected, mvarErrorNoRowsSelectedCount)
		iniReadErrorHandling("errorInvalidInput", mvarErrorInvalidInput, mvarErrorInvalidInputCount)
		iniReadErrorHandling("errorNoHost", mvarErrorNoHost, mvarErrorNoHostCount)
		iniReadErrorHandling("errorLoop", mvarErrorLoop, mvarErrorLoopCount)
		iniReadErrorHandling("errorRequery", mvarErrorRequery, mvarErrorRequeryCount)
		iniReadErrorHandling("errorRequeryYes", mvarErrorRequeryYes, mvarErrorRequeryYesCount)
		iniReadErrorHandling("errorRequeryNo", mvarErrorRequeryNo, mvarErrorRequeryNoCount)
		iniReadErrorHandling("errorTrialsExceeded", mvarErrorTrialsExceeded, mvarErrorTrialsExceededCount)
		iniReadErrorHandling("errorTrialsPossible", mvarErrorTrialsPossible, mvarErrorTrialsPossibleCount)
		iniReadErrorHandling("userOrderSuccess", mvarUserOrderSuccess, mvarUserOrderSuccessCount)
		iniReadErrorHandling("programOrderWritten", mvarUserOrderWritten, mvarUserOrderWrittenCount)
		iniReadErrorHandling("userOrderFailure", mvarUserOrderFailure, mvarUserOrderFailureCount)
		iniReadErrorHandling("programOrderRead", mvarUserOrderRead, mvarUserOrderReadCount)
		iniReadErrorHandling("userMotivSuccess", mvarUserMotivSuccess, mvarUserMotivSuccessCount)
		iniReadErrorHandling("programStart", mvarProgramStart, mvarProgramStartCount)
		iniReadErrorHandling("programEnd", mvarProgramEnd, mvarProgramEndCount)
		Dim tmp() As EVENTSTRUCT
		Dim i As Integer
		readDBEvents(tmp, i)
		ReDim mvarEventNames(mitaEventCodes.codesMax - 1)
		If Not IsNothing(tmp) Then
			For i = 0 To UBound(tmp)
				mvarEventNames(eventGetCodeValue(tmp(i).sName)) = tmp(i).sName
			Next
		End If
		Dim sb As New System.Text.StringBuilder(1000)
		sb.Append("SELECT COUNT(action) FROM " & mitaSystem.tableEventControl)
		sb.Append(" WHERE sapsystemid = " & mitaSystem.sapSystemId)
		sb.Append(" AND rfctype = " & mitaSystem.rfcType)
		sb.Append(" AND runtype = '" & mitaSystem.runType & "'")
		sb.Append(" AND action = 'DUMP'")
		sb.Append(" AND activ = 'Y'")
		sb.Append(" ORDER BY recno;")
		Dim query As String = sb.ToString()
		Dim result As Boolean
		Dim count As Integer
		result = mitaConnect.queryNumber(query, count)
		mvarNeedsLogs = result And (count > 0)
	End Sub

	'Public Function namesSetSapRfcType(ByRef sapRfcTyp As Integer) As Boolean
	'	mitaSystem.rfctype = sapRfcTyp
	'	Return True
	'End Function
	'Public Function namesSetConnectString(ByVal connectString As String) As String
	'	Dim errorMessage As String = dbTest(connectString)
	'	If errorMessage = "" Then
	'		mitaConnect.dbConnectString = connectString
	'		dbOpenOdbc()
	'		dbCloseOdbc()
	'		Return ""
	'	End If
	'	mitaData.frmMitaMsgInst.popupMessage(errorMessage, 0, "Logon")
	'	Return errorMessage
	'End Function
	'Public Function namesSetCommandLine(ByVal cmd As String) As Boolean
	'	mvarCommandLine = cmd
	'	Return True
	'End Function
	'Public Function namesSetSapVersion(ByRef sapVersion As Integer) As Boolean
	'	mitaData.sapSystemVERSIONID = sapVersion
	'	mvarSapError.namesSetSapVersion(sapVersion)
	'	namesSetSapVersion = True
	'End Function
	'Public Function namesSetSapSystem(ByRef sapSystem As Integer) As Boolean
	'	mitaData.sapSystemID = sapSystem
	'	namesSetSapSystem = True
	'	Return readDBSapSystemFromID(mitaData.sapSystemID)
	'End Function

	'Public Function namesSetSapSystemName(ByRef sapSystemNAME As String, ByRef sapVersionName As String) As Boolean
	'	Dim result As Boolean
	'	mitaData.sapName = sapSystemNAME
	'	namesSetSapSystemName = False
	'	result = readDBSapSystemFromName(mitaData.sapName)
	'	If result Then
	'		mitaData.sapSystemID = mitaData.sapSystemID
	'		sapVersionName = mitaData.sapSystemVERSIONNAME
	'		namesSetSapSystemName = True
	'	End If
	'End Function


	Public Function optionsSetSqlClass(ByRef classMap As Integer) As Boolean
		mvarLogClassMap = classMap
		optionsSetSqlClass = True
	End Function
	'Public Function optionsSetMiniSap(ByVal miniSap As Boolean) As Boolean
	'	mvarSapError.optionsSetMiniSap(miniSap)
	'	mitaData.isMiniSAP = miniSap
	'	Return True
	'End Function
	Public Function orderClose() As Boolean
		orderClose = False
		orderClose = True
		hasOrderBytes = False
		mvarOrderByteCount = 0
		ReDim mvarOrderBytes(0)
		mvarOrderOpen = 0
		mvarOrderInitialized = False
		mvarSelect = ""
	End Function

	Public Function orderAbort() As Boolean
		orderAbort = True
		hasOrderBytes = False
		mvarOrderByteCount = 0
		ReDim mvarOrderBytes(0)
		mvarOrderOpen = 0
		mvarOrderInitialized = False
	End Function
	Public Function getTableRecords(ByRef mitaIndex As Integer, ByRef targetRecords() As String) As Boolean
		Dim i As Integer
		getTableRecords = checkOpen()
		If Not getTableRecords Then Exit Function
		ReDim targetRecords(mvarSapRecords(mitaIndex).mtCount)
		For i = 0 To mvarSapRecords(mitaIndex).mtCount
			targetRecords(i) = mvarSapRecords(mitaIndex).mtRecords(i).sbContent
		Next i
	End Function

	'Public Function getParameterString(ByRef mitaIndex As Integer, ByRef targetParameter As String) As Boolean
	'	getParameterString = checkOpen()
	'	If Not getParameterString Then Exit Function
	'	targetParameter = mvarSapRecords(mitaIndex).mtRecords(0).sbContent
	'End Function

	Private Function getNextOrderId(ByRef id As Integer, Optional ByVal isError As Boolean = False) As Boolean
		Dim query As String
		getNextOrderId = False
		On Error GoTo isErr
		If Not isError Then
			query = "SELECT orderid" & mitaSystem.sapSystemId & ".nextval FROM dual"
		Else
			query = "SELECT errorid" & mitaSystem.sapSystemId & ".nextval FROM dual"
		End If
		Return mitaConnect.queryNumber(query, id)
Exx:
		On Error Resume Next
		Exit Function
isErr:
		eventProcess(Err.Description & vbCrLf & query, (mitaEventCodes.errorDatabaseSequence), "getNextOrderId")
		Resume Exx
	End Function

	'Private Function orderInitialize(ByRef RJHATPAK As String) As Boolean
	'	orderInitialize = False
	'	If Not checkInputs() Then Exit Function
	'	If mvarOrderOpen <> 0 Then orderAbort()
	'	resetAll()
	'	If Not mvarIniLoaded Then iniRead()
	'	mvarSapRecords(0).mtCount = 0
	'	mvarSapRecords(0).mtRecords(0).sbContent = RJHATPAK
	'	mvarSapRecords(0).mtRecords(0).sbUsed = True
	'	orderInitialize = True
	'	mvarOrderInitialized = True
	'End Function

	Private Function readByteLine(ByRef bytCnt As Integer, ByRef Lin As String) As Integer
		Lin = ""
		Do
			If bytCnt >= mvarOrderByteCount Then Exit Do
			If Chr(mvarOrderBytes(bytCnt)) = vbCr Then Exit Do
			Lin = Lin & Chr(mvarOrderBytes(bytCnt))
			bytCnt = bytCnt + 1
		Loop
		readByteLine = bytCnt + 2

	End Function
	Public Sub init()
		If Not mvarIniLoaded Then iniRead()
	End Sub
	Public Function orderReadNextDB() As Boolean
		Dim query As String
		'aTimer.Stop()
		If mvarOrderOpen <> 0 Then orderAbort()
		resetAll()
		If Not mvarIniLoaded Then iniRead()
		orderReadNextDB = False
		On Error GoTo isErr
		Dim psc_NextOrder As New OdbcCommand
		psc_NextOrder.Connection = mitaConnect.odbc_connection
		psc_NextOrder.CommandText = "{ CALL psc" & CStr(mitaSystem.sapSystemId) & "_nextorder_sp(?,?,?,?,?) }"
		psc_NextOrder.CommandType = CommandType.StoredProcedure
		psc_NextOrder.Parameters.Clear()
		psc_NextOrder.Parameters.Add("f_openid", OdbcType.Int).Value = mitaData.createID
		psc_NextOrder.Parameters.Add("f_sapsystemid", OdbcType.Int).Value = mitaSystem.sapSystemId
		psc_NextOrder.Parameters.Add("f_sapversionid", OdbcType.Int).Value = mitaSystem.sapSystemVERSIONID
		psc_NextOrder.Parameters.Add("f_pscid", OdbcType.Int).Direction = ParameterDirection.Output
		psc_NextOrder.Parameters.Add("f_processed", OdbcType.Int).Direction = ParameterDirection.Output
		mitaConnect.odbc_connection.Open()
		psc_NextOrder.ExecuteNonQuery()
		mitaConnect.odbc_connection.Close()
		mvarOrderID = CType(psc_NextOrder.Parameters("f_pscid").Value, Integer)
		If mvarOrderID <> -1 Then
			mvarNumberProcessed = CInt(psc_NextOrder.Parameters("f_processed").Value)
			orderReadNextDB = orderReadDB(mvarOrderID)
			mvarActualLength = mvarOrderByteCount
		End If
Exx:
		'aTimer.Start()
		Exit Function
isErr:
		mitaConnect.odbc_connection.Close()
		eventProcess(Err.Description & vbCrLf & query, (mitaEventCodes.errorProgrammer), "orderReadNextDB")
		Resume Exx
	End Function

	Private Sub iniReadErrorHandling(ByRef section As String, ByRef targetError() As errorStructure, ByRef count As Integer)
		Dim a As String
		Dim key As String
		Dim x() As String
		Dim i As Integer
		Dim errTmp(0) As EVENTSTRUCT
		ReDim targetError(0)
		count = 0
		readDBEventActions(section, errTmp, 0)
		count = errTmp(0).sCount
		If count = -1 Then
			ReDim targetError(0)
			Exit Sub
		End If
		ReDim targetError(count)
		For i = 0 To count
			targetError(i).eAction = errTmp(0).sActions(i).aName
			targetError(i).eParameter = errTmp(0).sActions(i).aTodo
		Next i
	End Sub

	Private Sub resetAll()
		Dim i As Integer
		For i = 1 To mitaData.structureCount
			mvarSapRecords(i).mtCount = 0
			ReDim mvarSapRecords(i).mtRecords(0)
		Next i
		mvarOrderOpen = 0
		hasOrderBytes = False
		mvarTpAdNo = ""
		mvarPaper = ""
		mvarTpAdVer = ""
		mvarMotivNo = CStr(0)
	End Sub
	Public Function selectAbort() As Boolean
		selectAbort = orderAbort()
	End Function


	Public Function selectExecuteQueryForReadOnly(ByRef whereClause As String) As Boolean
		Dim query As String
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		query = Trim(whereClause)
		If UCase(Left(query, 6)) = "WHERE " Then query = Trim(Mid(query, 7))
		mvarSelect = query
		If Not mvarIniLoaded Then iniRead()
		selectExecuteQueryForReadOnly = False
		mvarSelectCount = 0
		query = "SELECT pscid FROM " & mitaSystem.tableOrderControl & " WHERE " & mvarSelect
		query = query & " AND sapversionid = " & mitaSystem.sapSystemVERSIONID
		query = query & " AND sapsystemid = " & mitaSystem.sapSystemId
		idbc.CommandText = query
		Try
			mitaConnect.odbc_connection.Open()
			reader = idbc.ExecuteReader
			If reader.HasRows Then
				While reader.Read
					mvarSelectCount = mvarSelectCount + 1
					ReDim Preserve mvarSelectID(mvarSelectCount - 1)
					mvarSelectID(mvarSelectCount - 1) = CInt(reader.Item("pscid").ToString)
				End While
				selectExecuteQueryForReadOnly = True
			Else
				eventProcess("User Select gave no result rows", (mitaEventCodes.errorNoRowsSelected), "selectExecuteQueryForReadOnly")
			End If
			mvarSelectCursor = -1
			reader.Close()
			idbc.Dispose()
		Catch
			eventProcess(Err.Description & vbCrLf & query, (mitaEventCodes.errorProgrammer), "selectExecuteQueryForReadOnly")
			mvarSelect = ""
		End Try
		mitaConnect.odbc_connection.Close()
	End Function

	Public Function selectGetRowCount(ByRef targetRowCount As Integer) As Boolean
		selectGetRowCount = False
		If mvarSelect = "" Then
			eventProcess("No active select open", (mitaEventCodes.errorNoActualSelect), "selectGetRowCount")
			Exit Function
		End If
		targetRowCount = mvarSelectCount
		selectGetRowCount = True
	End Function


	Public Function selectReadNext() As Boolean
		selectReadNext = False
		If mvarSelect = "" Then
			eventProcess("No active select open", (mitaEventCodes.errorNoActualSelect), "selectReadNext")
			Exit Function
		End If
		If mvarSelectCount = 0 Then
			eventProcess("Select gave no result rows", (mitaEventCodes.errorNoRowsSelected), "selectReadNext")
			Exit Function
		End If
		If mvarSelectCursor = UBound(mvarSelectID) Then
			Exit Function
		End If
		mvarSelectCursor = mvarSelectCursor + 1
		selectReadNext = orderReadDB(mvarSelectID(mvarSelectCursor))
	End Function

	Public Function selectMoveFirst() As Boolean
		selectMoveFirst = False
		If mvarSelect = "" Then
			eventProcess("No active user select open", (mitaEventCodes.errorNoActualSelect), "selectMoveFirst")
			Exit Function
		End If
		If mvarSelectCount = 0 Then
			eventProcess("User Select gave no result rows", (mitaEventCodes.errorNoRowsSelected), "selectMoveFirst")
			Exit Function
		End If
		mvarSelectCursor = -1
		selectMoveFirst = True
	End Function

	Public Function optionsSetGarbageRemove(ByRef garbageRemove As Boolean) As Boolean
		mvarGarbageRemove = garbageRemove
		optionsSetGarbageRemove = True
	End Function

	Public Function optionsAllowOlderVersion(ByRef allowOlderVersion As Boolean) As Boolean
		mvarAllowOlderVersion = allowOlderVersion
		Return True
	End Function

	Public Function optionsAllowSameVersion(ByRef allowSameVersion As Boolean) As Boolean
		mvarAllowSameVersion = allowSameVersion
		Return True
	End Function

	Public Function optionsSaveAllSAP(ByRef saveAllSap As Boolean) As Boolean
		mvarSaveAllSAP = saveAllSap
		Return True
	End Function

	Private Sub iniRead()
		mvarMitaSqlIni = mvarParentPath & defaultSQL
		mvarIniIsLoading = True
		mvarIniLoaded = False
		'Dim sav As String = mitaData.runType
		'readDBSapSystemFromName(mitaData.sapSystemNAME)
		'If Not IsNothing(sav) Then mitaData.runType = sav
		If Not iniReadDB() Then
			RaiseEvent sqlError("Problems reading initialisation", "", "User Mistake")
			RaiseEvent endApplication(True)
			Exit Sub
		End If
		iniReadIntern()
		mvarIniIsLoading = False
		If mitaData.createHost = "" Then
			eventProcess("No HOSTNAME defined in environment!", (mitaEventCodes.errorNoHost), "Class_Initialize")
		End If
		mvarIniLoaded = True
		mvarActualControl = mitaSystem.tableOrderControl
		mvarActualData = mitaSystem.tableOrderData
	End Sub

	Private Function iniReadDB() As Boolean
		Dim i As Integer
		Dim j As Integer
		Dim k As Integer
		Dim buffer As String
		Dim a As String
		Dim p As Integer
		Dim tmpRec As MITAFIELD
		Dim result As Boolean
		Dim cFields() As FIELDSTRUCT
		Dim cCount As Integer
		iniReadDB = False
		mitaData.structureCount = -1
		result = readDBCustCombis(mitaData.combis, mitaData.combiCount)
		result = readDBStructures(mitaData.structures, mitaData.structureCount)
		If Not result Or mitaData.structureCount < 0 Then Exit Function
		ReDim mvarSapRecords(mitaData.structureCount)
		ReDim mvarSAPRecName(mitaData.structureCount)
		ReDim mvarSAPRecLen(mitaData.structureCount)
		ReDim mvarSAPRecord(mitaData.structureCount)
		For i = 0 To mitaData.structureCount
			mvarSapRecords(i).mtName = mitaData.structures(i).sName
			mvarSapRecords(i).mtLength = mitaData.structures(i).sLength
			mvarSapRecords(i).mtExt = "." & Mid(mvarSapRecords(i).mtName, StrRecord)
			mvarSAPRecName(i) = mitaData.structures(i).sName
			mvarSAPRecLen(i) = mitaData.structures(i).sLength
		Next i
		cCount = -1
		result = readDBCustFields(cFields, cCount)
		If Not result Or cCount < 0 Then Exit Function
		ReDim mvarSAP2DB(cCount + 1)
		For i = 0 To cCount
			tmpRec.mfFirst = CShort(cFields(i).fFirst)
			tmpRec.mfLabel = cFields(i).fName
			tmpRec.mfLength = CShort(cFields(i).fLength)
			tmpRec.mfTyp = cFields(i).fType
			For j = 0 To mitaData.structureCount
				If cFields(i).fStructure = mvarSapRecords(j).mtName Then
					tmpRec.mfRecord = j
					Exit For
				End If
			Next j
			If i = 0 Then
				mvarSAP2DB(i + 1) = tmpRec
			Else
				a = tmpRec.mfLabel
				For j = 1 To UBound(mvarSAP2DB) - 1
					If mvarSAP2DB(j).mfLabel > a Then
						For k = UBound(mvarSAP2DB) To j + 1 Step -1
							mvarSAP2DB(k) = mvarSAP2DB(k - 1)
						Next k
						mvarSAP2DB(j) = tmpRec
						Exit For
					ElseIf mvarSAP2DB(j).mfLabel = "" Then
						mvarSAP2DB(j) = tmpRec
						Exit For
					End If
				Next j
				If j = UBound(mvarSAP2DB) Then mvarSAP2DB(j) = tmpRec
			End If
		Next i
		cCount = -1
		result = readDBStructFields("RJHATSTAT", cFields, cCount)
		If Not result Or cCount < 0 Then Exit Function
		For i = 0 To cCount
			If UCase(Left(cFields(i).fName, 4)) = "MERK" Then
				mvarStatTableBegin = cFields(i).fFirst
				mvarStatLabelLength = cFields(i).fLength
				mvarStatContentLength = cFields(i + 1).fLength
				Exit For
			End If
		Next i
		result = readDBCustQueries(mvarCustomQueries, mvarCustomQueryNames)

		mvarFldProduct = findSapField(moINDEX, "Belegeinh")
		mvarFldOrderAVM = findSapField(pakINDEX, "AvmNr")
		mvarFldOrderMotiv = findSapField(moINDEX, "Motiv")
		mvarFldOrderVNO = findSapField(pakINDEX, "AenversNr")
		mvarRefPAPAVM = findSapField(papIndex, "RefAvmNr")
		mvarRefPAPPos = findSapField(papIndex, "RefPosNr")

		mvarFldPSEin = findSapField(psINDEX, "EinNr")
		mvarFldSTATEin = findSapField(statINDEX, "EinNr")
		mvarFldTXTEin = findSapField(txtINDEX, "EinNr")
		mvarFldPSIEin = findSapField(psiINDEX, "EinNr")
		mvarFldPLZEin = findSapField(plzINDEX, "EinNr")
		mvarFldPLZAEin = findSapField(plzaINDEX, "EinNr")

		mvarFldBPZPos = findSapField(bpzINDEX, "PosNr")
		mvarFldISZPos = findSapField(iszINDEX, "PosNr")
		mvarFldPAPPos = findSapField(papIndex, "PosNr")
		mvarFldPLZAPos = findSapField(plzaINDEX, "PosNr")
		mvarFldPLZPos = findSapField(plzINDEX, "PosNr")
		mvarFldPSPos = findSapField(psINDEX, "PosNr")
		mvarFldSTATPos = findSapField(statINDEX, "PosNr")
		mvarFldTXTPos = findSapField(txtINDEX, "PosNr")
		mvarFldPSIPos = findSapField(psiINDEX, "PosNr")

		mvarFldPSMo = findSapField(psINDEX, "Motiv")
		mvarFldMOMo = findSapField(moINDEX, "Motiv")
		mvarFldTXTMo = findSapField(txtINDEX, "Motiv")
		mvarFldSTATMo = findSapField(statINDEX, "Motiv")
		mvarFldPSIMo = findSapField(psiINDEX, "Motiv")
		mvarFldBLZMo = findSapField(blzINDEX, "Motiv")

		mvarFldSTATLevel = findSapField(statINDEX, "StatEbene")
		mvarFldTXTLevel = findSapField(txtINDEX, "TextEbene")

		If result Then iniReadDB = iniReadTables()
	End Function

	Private Function iniReadTables() As Boolean
		Dim i As Integer
		Dim tbl() As TABLESTRUCT
		If readDBCustTables(tbl, mvarCustTableCount) Then
			ReDim Preserve mvarCustTablesArray(mvarCustTableCount)
			For i = 0 To mvarCustTableCount - 1
				mvarCustTablesArray(i) = readDBCustTableEntries(tbl(i).tName)
				mvarCustTablesArray(i).name = tbl(i).tName
			Next i
			iniReadTables = True
		End If
	End Function

	Public Function orderWriteFile(ByRef fileName As String, ByRef fileSize As Integer) As Boolean
		Dim isOpen As Boolean
		Dim f1 As Integer
		orderWriteFile = False
		isOpen = False
		buildOrderBytes()
		If mvarOrderByteCount > 0 Then
			If Dir(fileName) <> "" Then Kill(fileName)
			On Error GoTo writeErr
			f1 = FreeFile()
			FileOpen(f1, fileName, OpenMode.Binary)
			isOpen = True
			FilePut(f1, mvarOrderBytes)
			FileClose(f1)
			orderWriteFile = True
			fileSize = mvarOrderByteCount
		End If
Exx:
		eventProcess("File", mitaEventCodes.programOrderWritten, "orderWriteFile")
		Exit Function
writeErr:
		If isOpen Then FileClose(1)
		eventProcess(Err.Description & vbCrLf & fileName, (mitaEventCodes.errorFileSystem), "orderWriteFile")
		Resume Exx
	End Function

	Private Sub buildOrderBytes()
		Dim a As String
		Dim p As Integer
		Dim f1 As Integer
		Dim i As Integer
		Dim rc As Integer
		Dim MyTable() As String
		Dim isOpen As Boolean
		mvarOrderByteCount = 0
		For i = LBound(mvarSapRecords) To UBound(mvarSapRecords)
			a = mvarSapRecords(i).mtName
			If a <> "" And mvarSapRecords(i).mtCount > 0 Then
				p = InStr(a, "_")
				If p <> 0 Then a = Left(a, p - 1)
				findRecords(a, MyTable)
				If MyTable(0) <> "" Then
					mvarOrderByteCount = mvarOrderByteCount + Len(a) + 4
					For p = 0 To UBound(MyTable)
						If mvarSapRecords(i).mtRecords(p).sbUsed Or Not mvarGarbageRemove Or i = psINDEX Then
							mvarOrderByteCount = mvarOrderByteCount + Len(MyTable(p)) + 2
						End If
					Next p
				End If
			End If
		Next i
		If mvarOrderByteCount = 0 Then Exit Sub
		ReDim mvarOrderBytes(mvarOrderByteCount - 3)
		mvarOrderByteCount = 0
		For i = LBound(mvarSapRecords) To UBound(mvarSapRecords)
			a = mvarSapRecords(i).mtName
			If a <> "" And mvarSapRecords(i).mtCount > 0 Then
				p = InStr(a, "_")
				If p <> 0 Then a = Left(a, p - 1)
				findRecords(a, MyTable)
				If MyTable(0) <> "" Then
					addBytes("[" & a & "]", ((i > 0) Or (p > 0)))
					For p = 0 To UBound(MyTable)
						If mvarSapRecords(i).mtRecords(p).sbUsed Or Not mvarGarbageRemove Or i = psINDEX Then
							addBytes(MyTable(p), True)
						End If
					Next p
				End If
			End If
		Next i
		hasOrderBytes = True
	End Sub

	Private Sub addBytes(ByRef inp As String, ByVal addCrLf As Boolean)
		Dim i As Integer
		If addCrLf Then
			mvarOrderBytes(mvarOrderByteCount) = CByte(Asc(vbCr))
			mvarOrderByteCount = mvarOrderByteCount + 1
			mvarOrderBytes(mvarOrderByteCount) = CByte(Asc(vbLf))
			mvarOrderByteCount = mvarOrderByteCount + 1
		End If
		For i = 1 To Len(inp)
			mvarOrderBytes(mvarOrderByteCount) = CByte(Asc(Mid(inp, i, 1)))
			mvarOrderByteCount = mvarOrderByteCount + 1
		Next i
	End Sub
	Private Function findSapField(ByVal struct As Integer, ByVal name As String) As MITAFIELD
		Dim result As Boolean
		Dim structName As String
		Dim cFields() As FIELDSTRUCT
		Dim cCount As Integer
		Dim i As Integer
		Dim res As New MITAFIELD
		Dim uName As String = name.ToUpper
		structName = mitaData.structures(struct).sName
		result = readDBStructFields(structName, cFields, cCount)
		res.mfRecord = struct
		For i = 0 To cCount
			If cFields(i).fName.ToUpper.Equals(uName) Then
				res.mfFirst = cFields(i).fFirst
				res.mfLength = cFields(i).fLength
				res.mfTyp = cFields(i).fType
				Exit For
			End If
		Next
		Return res
	End Function
	Private Sub findRecords(ByRef stct As String, ByRef MyTable() As String)
		Dim rc As Integer
		For rc = 0 To UBound(mvarSapRecords)
			If mvarSapRecords(rc).mtName = stct Then Exit For
		Next rc
		ReDim MyTable(0)
		Select Case rc
			Case pakINDEX
				MyTable = mvarSAPPakRec
			Case papIndex
				MyTable = mvarSAPPapRec
			Case bpzINDEX
				MyTable = mvarSAPBpzRec
			Case iszINDEX
				MyTable = mvarSAPIszRec
			Case moINDEX
				MyTable = mvarSAPMoRec
			Case blzINDEX
				MyTable = mvarSAPBlzRec
			Case psINDEX
				MyTable = mvarSAPPsRec
			Case plzINDEX
				MyTable = mvarSAPPlzRec
			Case plzaINDEX
				MyTable = mvarSAPPlzaRec
			Case statINDEX
				MyTable = mvarSAPStatRec
			Case txtINDEX
				MyTable = mvarSAPTxtRec
			Case psiINDEX
				MyTable = mvarSAPPsiRec
			Case Else
				MyTable(0) = ""
		End Select
	End Sub

	Private Function buildTree() As Boolean
		Dim tmpMotiv As MOTIV
		Dim m As Integer
		Dim p As Integer
		Dim i As Integer
		Dim cnt As Integer
		Dim result As Boolean
		Dim tst() As String
		result = False

		mvarSapRecords(pakINDEX).mtRecords(0).sbUsed = True
		mvarSapRecords(pakINDEX).mtCount = 1

		If statCount > -1 Then
			ReDim STAT(statCount)
			For p = 0 To statCount
				STAT(p).posNr = CInt(getSapValue(mvarFldSTATPos, mvarSAPStatRec(p)))
				STAT(p).einNr = CInt(getSapValue(mvarFldSTATEin, mvarSAPStatRec(p)))
				STAT(p).moNr = CInt(getSapValue(mvarFldSTATMo, mvarSAPStatRec(p)))
				STAT(p).level = CInt(getSapValue(mvarFldSTATLevel, mvarSAPStatRec(p)))
				STAT(p).used = False
			Next p
		End If
		If txtCount > -1 Then
			ReDim Txt(txtCount)
			For p = 0 To txtCount
				Txt(p).posNr = CInt(getSapValue(mvarFldTXTPos, mvarSAPTxtRec(p)))
				Txt(p).einNr = CInt(getSapValue(mvarFldTXTEin, mvarSAPTxtRec(p)))
				Txt(p).moNr = CInt(getSapValue(mvarFldTXTMo, mvarSAPTxtRec(p)))
				Txt(p).level = CInt(getSapValue(mvarFldTXTLevel, mvarSAPTxtRec(p)))
				Txt(p).used = False
			Next p
		End If

		If psiCount > -1 Then
			ReDim PSI(psiCount)
			For p = 0 To psiCount
				PSI(p).posNr = CInt(getSapValue(mvarFldPSIPos, mvarSAPPsiRec(p)))
				PSI(p).einNr = CInt(getSapValue(mvarFldPSIEin, mvarSAPPsiRec(p)))
				PSI(p).moNr = CInt(getSapValue(mvarFldPSIMo, mvarSAPPsiRec(p)))
				PSI(p).used = False
			Next p
		End If

		If moCount > -1 Then
			ReDim MO(moCount)
			For m = 0 To moCount
				MO(m).moNr = CInt(getSapValue(mvarFldMOMo, mvarSAPMoRec(m)))
				MO(m).used = False
				MO(m).statINDEX7 = -1
				MO(m).psiINDEX7 = -1
				MO(m).txtINDEX7 = -1
				MO(m).papCount = -1
			Next m
		End If

		If blzCount > -1 Then
			ReDim BLZ(blzCount)
			For m = 0 To blzCount
				BLZ(m).moNr = CInt(getSapValue(mvarFldBLZMo, mvarSAPBlzRec(m)))
				BLZ(m).used = False
			Next m
		End If

		If bpzCount > -1 Then
			ReDim BPZ(bpzCount)
			For p = 0 To bpzCount
				BPZ(p).posNr = CInt(getSapValue(mvarFldBPZPos, mvarSAPBpzRec(p)))
				BPZ(p).used = False
			Next p
		End If

		If iszCount > -1 Then
			ReDim ISZ(iszCount)
			For p = 0 To iszCount
				ISZ(p).posNr = CInt(getSapValue(mvarFldISZPos, mvarSAPIszRec(p)))
				ISZ(p).used = False
			Next p
		End If

		If plzCount > -1 Then
			ReDim PLZ(plzCount)
			For p = 0 To plzCount
				PLZ(p).posNr = CInt(getSapValue(mvarFldPLZPos, mvarSAPPlzRec(p)))
				PLZ(p).einNr = CInt(getSapValue(mvarFldPLZEin, mvarSAPPlzRec(p)))
				PLZ(p).used = False
			Next p
		End If

		If plzaCount > -1 Then
			ReDim PLZA(plzCount)
			For p = 0 To plzaCount
				PLZA(p).posNr = CInt(getSapValue(mvarFldPLZAPos, mvarSAPPlzaRec(p)))
				PLZA(p).einNr = CInt(getSapValue(mvarFldPLZAEin, mvarSAPPlzaRec(p)))
				PLZA(p).used = False
			Next p
		End If

		If papCount > -1 Then
			ReDim PAP(papCount)
			For p = 0 To papCount
				PAP(p).posNr = CInt(getSapValue(mvarFldPAPPos, mvarSAPPapRec(p)))
				PAP(p).used = False
				PAP(p).psiINDEX3 = -1
				PAP(p).statINDEX3 = -1
				PAP(p).txtINDEX3 = -1
			Next p
		End If

		If psCount > -1 Then
			ReDim PS(psCount)
			For i = 0 To psCount
				PS(i).posNr = CInt(getSapValue(mvarFldPSPos, mvarSAPPsRec(i)))
				PS(i).einNr = CInt(getSapValue(mvarFldPSEin, mvarSAPPsRec(i)))
				PS(i).moNr = CInt(getSapValue(mvarFldPSMo, mvarSAPPsRec(i)))
				PS(i).papIndex = -1
				PS(i).moINDEX = -1
				PS(i).blzINDEX = -1
				PS(i).plzINDEX = -1
				PS(i).plzaINDEX = -1
				PS(i).statINDEX8 = -1
				PS(i).txtINDEX8 = -1
				PS(i).psiINDEX8 = -1
				PS(i).used = True
				mvarSapRecords(psINDEX).mtRecords(i).sbUsed = True

				For m = 0 To blzCount
					If BLZ(m).moNr = PS(i).moNr Then
						PS(i).blzINDEX = i
						BLZ(m).used = True
						mvarSapRecords(blzINDEX).mtRecords(m).sbUsed = True
						Exit For
					End If
				Next m
				'suchen nach pos
				For m = 0 To bpzCount
					If BPZ(m).posNr = PS(i).posNr Then
						BPZ(m).used = True
						mvarSapRecords(bpzINDEX).mtRecords(m).sbUsed = True
						Exit For
					End If
				Next m

				For m = 0 To iszCount
					If ISZ(m).posNr = PS(i).posNr Then
						ISZ(m).used = True
						mvarSapRecords(iszINDEX).mtRecords(m).sbUsed = True
						Exit For
					End If
				Next m

				For p = 0 To papCount
					If PAP(p).posNr = PS(i).posNr Then
						PS(i).papIndex = m
						PAP(p).used = True
						mvarSapRecords(papIndex).mtRecords(p).sbUsed = True
						' suchen nach motiv
						For m = 0 To moCount
							If MO(m).moNr = PS(i).moNr Then
								PS(i).moINDEX = m
								MO(m).used = True
								Dim found As Integer = False
								Dim s As Integer
								For s = 0 To MO(m).papCount
									If MO(m).papIndexes(s) = p Then
										found = True
										Exit For
									End If
								Next
								If Not found Then
									MO(m).papCount = MO(m).papCount + 1
									ReDim Preserve MO(m).papIndexes(MO(m).papCount)
									MO(m).papIndexes(MO(m).papCount) = p
									PAP(p).psCount = -1
								End If
								PAP(p).psCount = PAP(p).psCount + 1
								ReDim Preserve PAP(p).psIndexes(PAP(p).psCount)
								PAP(p).psIndexes(PAP(p).psCount) = i
								mvarSapRecords(moINDEX).mtRecords(m).sbUsed = True
								Exit For
							End If
						Next m
						Exit For
					End If
				Next p

				'suchen nach ein und pos
				For m = 0 To plzCount
					If PLZ(m).posNr = PS(i).posNr And PLZ(m).einNr = PS(i).einNr Then
						PS(i).plzINDEX = m
						PLZ(m).used = True
						mvarSapRecords(plzINDEX).mtRecords(m).sbUsed = True
						Exit For
					End If
				Next m

				For m = 0 To plzaCount
					If PLZA(m).posNr = PS(i).posNr And PLZA(m).einNr = PS(i).einNr Then
						PS(i).plzaINDEX = m
						PLZA(m).used = True
						mvarSapRecords(plzaINDEX).mtRecords(m).sbUsed = True
						Exit For
					End If
				Next m

				For m = 0 To statCount
					If STAT(m).posNr = PS(i).posNr And STAT(m).einNr = PS(i).einNr And STAT(m).level = 8 Then
						PS(i).statINDEX8 = m
						STAT(m).used = True
						mvarSapRecords(statINDEX).mtRecords(m).sbUsed = True
						Exit For
					End If
				Next m

				For m = 0 To txtCount
					If Txt(m).posNr = PS(i).posNr And Txt(m).einNr = PS(i).einNr And Txt(m).level = 8 Then
						PS(i).txtINDEX8 = m
						Txt(m).used = True
						mvarSapRecords(txtINDEX).mtRecords(m).sbUsed = True
						Exit For
					End If
				Next m

				For m = 0 To psiCount
					If PSI(m).posNr = PS(i).posNr And PSI(m).einNr = PS(i).einNr And PSI(m).level = 8 Then
						PS(i).psiINDEX8 = m
						PSI(m).used = True
						mvarSapRecords(psiINDEX).mtRecords(m).sbUsed = True
						Exit For
					End If
				Next m
			Next i
		End If

		For m = 0 To moCount
			If MO(m).used Then
				For p = 0 To statCount
					If STAT(p).moNr = MO(m).moNr And STAT(p).level = 7 And MO(m).used Then
						MO(m).statINDEX7 = p
						STAT(p).used = True
						mvarSapRecords(statINDEX).mtRecords(p).sbUsed = True
						Exit For
					End If
				Next p
				For p = 0 To txtCount
					If Txt(p).moNr = MO(m).moNr And Txt(p).level = 7 And MO(m).used Then
						MO(m).txtINDEX7 = p
						Txt(p).used = True
						mvarSapRecords(txtINDEX).mtRecords(p).sbUsed = True
						Exit For
					End If
				Next p
				For p = 0 To psiCount
					If PSI(p).moNr = MO(m).moNr And PSI(p).level = 7 And MO(m).used Then
						MO(m).psiINDEX7 = p
						PSI(p).used = True
						mvarSapRecords(psiINDEX).mtRecords(p).sbUsed = True
						Exit For
					End If
				Next p
			End If
		Next m

		For p = 0 To papCount
			If PAP(p).used Then
				For m = 0 To statCount
					If STAT(m).posNr = PAP(p).posNr And STAT(m).level = 3 And PAP(p).used Then
						PAP(p).statINDEX3 = m
						STAT(m).used = True
						mvarSapRecords(statINDEX).mtRecords(m).sbUsed = True
						Exit For
					End If
				Next m
				For m = 0 To txtCount
					If Txt(m).posNr = PAP(p).posNr And Txt(m).level = 3 And PAP(p).used Then
						PAP(p).txtINDEX3 = m
						Txt(m).used = True
						mvarSapRecords(txtINDEX).mtRecords(m).sbUsed = True
						Exit For
					End If
				Next m
				For m = 0 To psiCount
					If PSI(m).posNr = PAP(p).posNr And PSI(m).level = 3 And PAP(p).used Then
						PAP(p).psiINDEX3 = m
						PSI(m).used = True
						mvarSapRecords(psiINDEX).mtRecords(m).sbUsed = True
						Exit For
					End If
				Next m
			End If
		Next p

		mvarOrderNo = getSapValue(mvarFldOrderAVM, mvarSAPPakRec(0))
		mvarVNO = CStr(CInt(getSapValue(mvarFldOrderVNO, mvarSAPPakRec(0))))
		'mvarOrderNo = Left(mvarSAPPakRec(0), 10)
		'mvarVNO = CStr(CInt(Mid(mvarSAPPakRec(0), 11, 4)))
		createVersionInfo()
		If mvarOrderNo = "" Then
			eventProcess("$AVM$($VNO$), RJHATPAK", (mitaEventCodes.errorSAPData), "buildTree")
			If Not mvarSaveAllSAP Then
				eventProcess("", mitaEventCodes.userOrderFailure, "buildTree")
				mvarOrderError = mitaErrorCodes.orderSapDataError
				Exit Function
			End If
		End If
		If moCount = -1 Then
			eventProcess("$AVM$($VNO$), RJHATMO", (mitaEventCodes.errorSAPData), "buildTree")
			If Not mvarSaveAllSAP Then
				eventProcess("", mitaEventCodes.userOrderFailure, "buildTree")
				mvarOrderError = mitaErrorCodes.orderSapDataError
				Exit Function
			End If
			result = True
			mvarOrderOpen = 1
		Else
			orderResetMotiv()
			If orderNextAllMotiv() Then
				result = True
			Else
				eventProcess("$AVM$($VNO$), RJHATPS", (mitaEventCodes.errorSAPData), "buildTree")
				If Not mvarSaveAllSAP Then
					eventProcess("", mitaEventCodes.userOrderFailure, "buildTree")
					mvarOrderError = mitaErrorCodes.orderNoPublishDate
					Exit Function
				Else
					result = True
				End If
			End If
		End If
		buildTree = result
	End Function

	'Public Function itemGetOrderNo(ByRef targetString As String) As Boolean
	'	itemGetOrderNo = checkOpen()
	'	If Not itemGetOrderNo Then Exit Function
	'	targetString = mvarOrderNo
	'End Function
	Public ReadOnly Property valueOrderString() As String
		Get
			Return mvarOrderNo
		End Get
	End Property
	'Public Function itemGetSapNo(ByRef targetString As String) As Boolean
	'	itemGetSapNo = checkOpen()
	'	If Not itemGetSapNo Then Exit Function
	'	targetString = mvarOrderNo & mvarMotivNo
	'End Function
	Public ReadOnly Property valueSapString() As String
		Get
			Return mvarOrderNo & mvarMotivNo
		End Get
	End Property
	'Public Function itemGetComboNo(ByRef targetValue As Integer) As Boolean
	'	itemGetComboNo = checkOpen()
	'	If Not itemGetComboNo Then Exit Function
	'	targetValue = comboIndex
	'End Function
	Public ReadOnly Property valueComboNo() As Integer
		Get
			Return comboIndex
		End Get
	End Property

	'Public Function itemGetMotivNo(ByRef targetString As String) As Boolean
	'	itemGetMotivNo = checkOpen()
	'	If Not itemGetMotivNo Then Exit Function
	'	targetString = mvarMotivNo
	'End Function
	Public ReadOnly Property valueMotivString() As String
		Get
			Return mvarMotivNo
		End Get
	End Property
	Public Function itemTestReference() As Boolean
		itemTestReference = checkOpen()
		If Not itemTestReference Then Exit Function
		itemTestReference = (mvarRefAVM <> "")
	End Function

	'Public Function getSapNo(ByRef targetSapNo As String) As Boolean
	'	getSapNo = checkOpen()
	'	If Not getSapNo Then Exit Function
	'	targetSapNo = mvarSapNo
	'End Function

	'Public Function itemGetOrderVNO(ByRef targetInteger As Integer) As Boolean
	'	itemGetOrderVNO = checkOpen()
	'	If Not itemGetOrderVNO Then Exit Function
	'	targetInteger = CInt(mvarVNO)
	'End Function
	Public ReadOnly Property valueOrderVno() As Integer
		Get
			Return CInt(mvarVNO)
		End Get
	End Property
	'Public Function itemGetPosNo(ByRef targetInteger As Integer) As Boolean
	'	itemGetPosNo = checkOpen() And CInt(mvarPosNo) > 0
	'	If Not itemGetPosNo Then Exit Function
	'	targetInteger = CInt(mvarPosNo)
	'End Function
	Public ReadOnly Property valuePosNo() As Integer
		Get
			Return CInt(mvarPosNo)
		End Get
	End Property

	'Public Function itemGetEinNo(ByRef targetInteger As Integer) As Boolean
	'	itemGetEinNo = checkOpen() And CInt(mvarEinNo) > 0
	'	If Not itemGetEinNo Then Exit Function
	'	targetInteger = CInt(mvarEinNo)
	'End Function


	Private Sub searchField(ByRef fieldName As String)
		Dim von As Integer
		Dim bis As Integer
		Dim test As Integer
		von = LBound(mvarSAP2DB)
		bis = UBound(mvarSAP2DB)
		Do
			If mvarSAP2DB(von).mfLabel = fieldName Then
				mvarActField = mvarSAP2DB(von)
				Exit Sub
			End If
			If mvarSAP2DB(bis).mfLabel = fieldName Then
				mvarActField = mvarSAP2DB(bis)
				Exit Sub
			End If
			If bis - von <= 1 Then Exit Sub
			test = CInt((bis - von) / 2 + von)
			If mvarSAP2DB(test).mfLabel = fieldName Then
				mvarActField = mvarSAP2DB(test)
				Exit Sub
			End If
			If mvarSAP2DB(test).mfLabel > fieldName Then
				bis = test
			Else
				von = test
			End If
		Loop
	End Sub

	Private Function searchQuery(ByRef customName As String) As Integer
		Dim von As Integer
		Dim bis As Integer
		Dim test As Integer
		Dim queryName As String
		von = LBound(mvarCustomQueryNames)
		bis = UBound(mvarCustomQueryNames)
		queryName = UCase(customName)
		searchQuery = -1
		Do
			If mvarCustomQueryNames(von) = queryName Then
				searchQuery = von
				Exit Function
			End If
			If mvarCustomQueryNames(bis) = queryName Then
				searchQuery = bis
				Exit Function
			End If
			If bis - von <= 1 Then
				Exit Function
			End If
			test = CInt((bis - von) / 2 + von)
			If mvarCustomQueryNames(test) = queryName Then
				searchQuery = test
				Exit Function
			End If
			If mvarCustomQueryNames(test) > queryName Then
				bis = test
			Else
				von = test
			End If
		Loop
	End Function

	Public Function orderWriteDB(ByRef blobSize As Integer) As Boolean
		Dim query As String
		Dim activ As String
		Dim result As Boolean
		Dim actVer As Integer
		Dim transact As OdbcTransaction
		Dim idbc As OdbcCommand
		blobSize = 0
		orderWriteDB = False
		If Not checkOpen() Then Exit Function
		buildOrderBytes()
		If Not mvarIsNew Then
			If mvarIsOlderVersion And Not mvarAllowOlderVersion Then
				eventProcess("$AVM$($VNO$)", mitaEventCodes.errorSAPVersion, "orderWriteDB")
				mvarOrderError = mitaErrorCodes.orderSapOldVersion
				Exit Function
			End If
			If mvarIsSameVersion And Not mvarAllowSameVersion Then
				eventProcess("$AVM$($VNO$)", mitaEventCodes.errorSAPVersion, "orderWriteDB")
				mvarOrderError = mitaErrorCodes.orderSapOldVersion
				Exit Function
			End If
		End If
		mvarActualLength = mvarOrderByteCount
		query = sqlTrans(mvarGetByVersion)
		If Not getNextOrderId(mvarOrderID) Then Exit Function
		mitaConnect.odbc_connection.Open()
		idbc = mitaConnect.odbc_connection.CreateCommand()
		transact = mitaConnect.odbc_connection.BeginTransaction()
		idbc.Transaction = transact
		query = mvarNewData
		result = sqlExecTrans(query, idbc)
		If result = False Then
			transact.Rollback()
			mitaConnect.odbc_connection.Close()
			Exit Function
		End If
		On Error GoTo isErr
		'buildHashCode()
		Dim i As Integer
		Dim a As Integer
		For i = 0 To hashCodeCount
			a = (i Xor mvarOrderBytes(i)) Mod 256
			mvarHashCode(i) = CByte(a)
		Next i
		query = sqlTrans(mvarWriteBlobs)
		idbc.CommandText = query
		idbc.Parameters.Add("hashcode", Odbc.OdbcType.VarBinary).Value = mvarHashCode
		idbc.Parameters.Add("structures", Odbc.OdbcType.VarBinary).Value = mvarOrderBytes
		idbc.ExecuteNonQuery()
		query = mvarNewControl
		result = sqlExecTrans(query, idbc)
		If result Then
			If mvarIsRetry Then
				query = mvarRetryControl
			Else
				query = mvarUpdateControl
			End If
			result = sqlExecTrans(query, idbc)
		End If
		If result And Not mvarAllowOlderVersion Then
			result = sqlExecTrans(mvarDeactivate, idbc)
		End If
		If result = False Then
			mvarOrderError = mitaErrorCodes.orderWriteDBProblem
			transact.Rollback()
		Else
			transact.Commit()
			If Not IsNothing(blobSize) Then blobSize = mvarOrderByteCount
			orderWriteDB = True
		End If
Exx:
		On Error Resume Next
		mitaConnect.odbc_connection.Close()
		If mvarOrderError = mitaErrorCodes.orderOK Then
			eventProcess("DB", mitaEventCodes.programOrderWritten, "orderWriteDB")
		End If
		Exit Function
isErr:
		transact.Rollback()
		eventProcess(Err.Description & vbCrLf & query, (mitaEventCodes.errorProgrammer), "orderWriteDB")
		Resume Exx
	End Function

	Public Function errorWriteDB(ByRef blobSize As Integer) As Boolean
		Dim query As String
		Dim activ As String
		Dim result As Boolean
		Dim actVer As Integer
		Dim transact As OdbcTransaction
		Dim idbc As OdbcCommand
		blobSize = 0
		errorWriteDB = False
		'If Not checkOpen() Then Exit Function
		'buildOrderBytes()
		Dim sav As Integer = mvarOrderID
		If Not getNextOrderId(mvarOrderID, True) Then GoTo exx
		mvarActualControl = mitaSystem.tableErrorControl
		mvarActualData = mitaSystem.tableErrorData
		mvarActualLength = mvarErrorByteCount
		mitaConnect.odbc_connection.Open()
		idbc = mitaConnect.odbc_connection.CreateCommand()
		transact = mitaConnect.odbc_connection.BeginTransaction()
		idbc.Transaction = transact
		query = mvarNewData
		result = sqlExecTrans(query, idbc)
		If result = False Then
			transact.Rollback()
			GoTo exx
		End If
		On Error GoTo isErr
		'buildHashCode()
		Dim i As Integer
		Dim a As Integer
		For i = 0 To hashCodeCount
			a = (i Xor mvarOrderBytes(i)) Mod 256
			mvarHashCode(i) = CByte(a)
		Next i
		query = sqlTrans(mvarWriteBlobs)
		idbc.CommandText = query
		idbc.Parameters.Add("hashcode", Odbc.OdbcType.VarBinary).Value = mvarHashCode
		idbc.Parameters.Add("structures", Odbc.OdbcType.VarBinary).Value = mvarErrorBytes
		idbc.ExecuteNonQuery()

		query = mvarNewControl
		result = sqlExecTrans(query, idbc)
		'If result Then
		'	If mvarIsRetry Then
		'		query = mvarRetryControl
		'	Else
		'		query = mvarUpdateControl
		'	End If
		'	result = sqlExecTrans(query, idbc)
		'End If
		'If result And Not mvarAllowOlderVersion Then
		'	result = sqlExecTrans(mvarDeactivate, idbc)
		'End If
		If result = False Then
			mvarOrderError = mitaErrorCodes.orderWriteDBProblem
			transact.Rollback()
		Else
			transact.Commit()
			If Not IsNothing(blobSize) Then blobSize = mvarErrorByteCount
			errorWriteDB = True
		End If
Exx:
		On Error Resume Next
		mitaConnect.odbc_connection.Close()
		'If mvarOrderError = mitaErrorCodes.orderOK Then
		'	eventProcess("DB", mitaEventCodes.programOrderWritten, "errorWriteDB")
		'End If
		mvarOrderID = sav
		mvarActualControl = mitaSystem.tableOrderControl
		mvarActualData = mitaSystem.tableOrderData
		mvarActualLength = mvarOrderByteCount
		Exit Function
isErr:
		transact.Rollback()
		eventProcess(Err.Description & vbCrLf & query, (mitaEventCodes.errorProgrammer), "orderWriteDB")
		Resume Exx
	End Function
	Private Sub buildHashCode()
		Dim i As Integer
		Dim a As Integer
		For i = 0 To hashCodeCount
			a = (i Xor mvarOrderBytes(i)) Mod 256
			mvarHashCode(i) = CByte(a)
		Next i
	End Sub

	Private Function orderWriteDBLog() As Boolean
		Dim query As String
		Dim activ As String
		Dim result As Boolean
		Dim isNew As Boolean
		orderWriteDBLog = False
		query = mvarNewLog
		result = sqlExecTrans(query)
		mvarActTime = Now

		orderWriteDBLog = True
		mvarLogCount = mvarLogCount + 1
Exx:
		Exit Function
isErr:
		eventProcess(Err.Description & vbCrLf & query, (mitaEventCodes.errorProgrammer), "orderWriteDB")
		Resume Exx
	End Function


	Private Function sqlExec(ByRef query As String, Optional ByRef cmd As OdbcCommand = Nothing) As Boolean
		Dim savStatus As ConnectionState = mitaConnect.odbc_connection.State
		sqlExec = False

		Dim idbc As OdbcCommand
		If IsNothing(cmd) Then
			idbc = mitaConnect.odbc_connection.CreateCommand()
		Else
			idbc = cmd
		End If
		idbc.CommandText = query
		Try
			If savStatus = ConnectionState.Closed Then mitaConnect.odbc_connection.Open()
			idbc.ExecuteNonQuery()
			If IsNothing(cmd) Then idbc.Dispose()
			sqlExec = True
		Catch
			eventProcess(Err.Description & vbCrLf & query, (mitaEventCodes.errorProgrammer), "sqlExec")
		End Try
		If IsNothing(cmd) Then If savStatus = ConnectionState.Closed Then mitaConnect.odbc_connection.Close()
	End Function


	Private Function sqlExecTrans(ByRef statement As String, Optional ByRef cmd As OdbcCommand = Nothing) As Boolean
		sqlExecTrans = sqlExec(sqlTrans(statement), cmd)
	End Function

	Private Function sqlTrans(ByRef statement As String) As String
		Dim sb As New System.Text.StringBuilder(statement, 1000)
		sb.Replace("#ID#", CStr(mvarOrderID))
		If mvarOrderNo = "" Then
			sb.Replace("'#AVM#'", "NULL")
			sb.Replace("#VNO#", "NULL")
		Else
			sb.Replace("#AVM#", mvarOrderNo)
			sb.Replace("#VNO#", CStr(mvarVNO))
		End If
		sb.Replace("#LENGTH#", CStr(mvarActualLength))
		sb.Replace("#PROCESSED#", CStr(mvarNumberProcessed))
		sb.Replace("#ORDERID#", CStr(mvarOrderID))
		sb.Replace("#MYSTATUS#", mvarMyStatus)
		sb.Replace("#EVENT#", mvarEventName)
		sb.Replace("#SAPVERSION#", CStr(mitaSystem.sapSystemVERSIONID))
		sb.Replace("#SAPSYSTEM#", CStr(mitaSystem.sapSystemId))
		sb.Replace("#PRIORITY#", CStr(mvarPriority))
		sb.Replace("#LOGTYP#", mvarLogTyp)
		sb.Replace("#TEXT#", mvarText)
		sb.Replace("#APP#", CStr(mitaSystem.rfcType))
		sb.Replace("#LOGCOUNT#", CStr(mvarLogCount))
		sb.Replace("#CONTROL#", CStr(mvarActualControl))
		sb.Replace("#LOG#", CStr(mitaSystem.tableEventLog))
		sb.Replace("#DATA#", CStr(mvarActualData))
		If InStr(statement, "#SYSTIME#") <> 0 Then
			sb.Replace("#SYSTIME#", CStr(Now))
		End If
		If mvarTpAdNo = "" Then
			sb.Replace("#ADNO#", "NULL")
			sb.Replace("#ADVER#", "NULL")
			sb.Replace("'#PAPER#'", "NULL")
		Else
			sb.Replace("#ADNO#", mvarTpAdNo)
			If mvarPaper = "" Then
				sb.Replace("'#PAPER#'", "NULL")
			Else
				sb.Replace("#PAPER#", mvarPaper)
			End If
			If mvarTpAdVer = "" Then
				sb.Replace("#ADVER#", "NULL")
			Else
				sb.Replace("#ADVER#", mvarTpAdVer)
			End If
		End If
		If mvarComment = "" Then
			sb.Replace("'#COMMENT#'", "NULL")
		Else
			If Len(mvarComment) > 253 Then
				sb.Replace("#COMMENT#", Left$(mvarComment, 253))
			Else
				sb.Replace("#COMMENT#", mvarComment)
			End If
		End If
		sb.Replace("#MYID#", CStr(mitaData.createID))
		If mitaData.createHost = "" Then
			sb.Replace("'#HOST#'", "NULL")
		Else
			sb.Replace("#HOST#", mitaData.createHost)
		End If
		If mitaData.createUser = "" Then
			sb.Replace("'#USER#'", "NULL")
		Else
			sb.Replace("#USER#", mitaData.createUser)
		End If
		If InStr(statement, "#LASTERROR#") > 0 Then
			If Len(mvarLastError) > 253 Then
				sb.Replace("#COMMENT#", Left$(mvarLastError, 253))
			Else
				sb.Replace("#LASTERROR#", mvarLastError)
			End If
		End If
		If InStr(statement, "#MOTIV#") > 0 Then
			If mvarMotivNo = "" Then
				sb.Replace("#MOTIV#", "NULL")
			ElseIf CShort(mvarMotivNo) = 0 Then
				sb.Replace("#MOTIV#", "NULL")
			Else
				sb.Replace("#MOTIV#", CStr(CInt(mvarMotivNo)))
			End If
		End If
		Return sb.ToString
	End Function
	Private Function orderReadDB(ByRef orderID As Integer) As Boolean
		Dim query As String
		mvarStartReadTime = VB.Timer()
		mvarStartOrderTime = VB.Timer()
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		If mvarOrderOpen <> 0 Then orderAbort()
		resetAll()
		mvarOrderID = orderID
		If Not mvarIniLoaded Then iniRead()
		If Not checkInputs() Then Exit Function
		orderReadDB = False
		query = sqlTrans(mvarFind)
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		If reader.Read Then
			mvarOrderByteCount = reader.GetInt32(0)
			mvarActualLength = mvarOrderByteCount
			ReDim mvarOrderBytes(mvarOrderByteCount - 1)
			reader.GetBytes(1, 0, mvarOrderBytes, 0, mvarOrderByteCount)
			mvarOrderNo = reader.GetString(2)
			mvarVNO = CStr(reader.GetInt32(3))
			reader.Close()
			mvarOrderOpen = 1
			If readOrderBytes() Then
				orderReadDB = True
				mvarOrderOpen = -1
			End If
		Else
			reader.Close()
		End If
		mitaConnect.odbc_connection.Close()
		If mvarOrderOpen <> 0 Then orderResetMotiv()
		eventProcess("", mitaEventCodes.programOrderRead, "orderReadDB")
	End Function

	Public Function orderReadFile(ByRef fileName As String) As Boolean
		Dim f1 As Integer
		Dim nam As String
		Dim errcnt As Integer
		mvarStartReadTime = VB.Timer()
		mvarStartOrderTime = VB.Timer()
		'aTimer.Stop()
		If mvarOrderOpen <> 0 Then orderAbort()
		resetAll()
		If Not mvarIniLoaded Then iniRead()
		If Not checkInputs() Then Exit Function
		nam = fileName
		f1 = FreeFile()
		orderReadFile = False
		On Error GoTo err_loop
		mvarOrderByteCount = CInt(FileLen(nam))
		mvarActualLength = mvarOrderByteCount
		ReDim mvarOrderBytes(mvarOrderByteCount - 1)
		FileOpen(f1, nam, OpenMode.Binary)
		FileGet(f1, mvarOrderBytes)
		On Error GoTo 0
		mvarOrderOpen = 1
		orderReadFile = readOrderBytes()
		hasOrderBytes = orderReadFile
ende:
		On Error Resume Next
		FileClose(f1)
		'aTimer.Start()
		If hasOrderBytes Then orderResetMotiv()
		eventProcess("", mitaEventCodes.programOrderRead, "orderReadDB")
		Exit Function
err_loop:
		errcnt = errcnt + 1
		If errcnt > 20 Then
			orderReadFile = False
			eventProcess(Err.Description & vbCrLf & nam, (mitaEventCodes.errorFileSystem), "orderReadFile")
			Resume ende
		Else
			Pause(2)
			Resume
		End If
	End Function

	Private Function readOrderBytes() As Boolean
		Dim Lin As String
		Dim nam As String
		Dim a As String
		Dim hasOld As Integer
		Dim errcnt As Integer
		Dim m As Integer
		Dim p As Integer
		Dim i As Integer
		Dim bytCnt As Integer
		mvarOrderError = mitaErrorCodes.orderOK
		hasOld = -1
		bytCnt = 0
		p = 0
		Do
			If bytCnt >= mvarOrderByteCount Then
				mvarSapRecords(hasOld).mtCount = p
				Exit Do
			End If
			bytCnt = readByteLine(bytCnt, Lin)
			If Left(Lin, 1) <> " " Then
				If Left(Lin, 1) = "[" Then
					Lin = Mid(Lin, 2, Len(Lin) - 2)
					For i = 0 To mitaData.structureCount
						If Lin = mvarSapRecords(i).mtName Then Exit For
					Next i
					ReDim mvarSapRecords(i).mtRecords(0)
					If hasOld <> -1 Then
						mvarSapRecords(hasOld).mtCount = p
					End If
					p = 0
					hasOld = i
				Else
					ReDim Preserve mvarSapRecords(i).mtRecords(p)
					mvarSapRecords(i).mtRecords(p).sbContent = Lin
					mvarSapRecords(i).mtRecords(p).sbUsed = False
					p = p + 1
				End If
			End If
		Loop
		pakCount = -1
		papCount = -1
		bpzCount = -1
		iszCount = -1
		moCount = -1
		blzCount = -1
		psCount = -1
		plzCount = -1
		plzaCount = -1
		statCount = -1
		txtCount = -1
		psiCount = -1
		For i = 0 To mitaData.structureCount
			Select Case i
				Case pakINDEX
					mvarSAPPakRec(pakINDEX) = mvarSapRecords(pakINDEX).mtRecords(0).sbContent
					mvarSAPRecord(0) = mvarSAPPakRec(pakINDEX)
					pakCount = 0
				Case papIndex : papCount = structFillRecs(i, mvarSAPPapRec)
				Case bpzINDEX : bpzCount = structFillRecs(i, mvarSAPBpzRec)
				Case iszINDEX : iszCount = structFillRecs(i, mvarSAPIszRec)
				Case moINDEX : moCount = structFillRecs(i, mvarSAPMoRec)
				Case blzINDEX : blzCount = structFillRecs(i, mvarSAPBlzRec)
				Case psINDEX : psCount = structFillRecs(i, mvarSAPPsRec)
				Case plzINDEX : plzCount = structFillRecs(i, mvarSAPPlzRec)
				Case plzaINDEX : plzaCount = structFillRecs(i, mvarSAPPlzaRec)
				Case statINDEX : statCount = structFillRecs(i, mvarSAPStatRec)
				Case txtINDEX : txtCount = structFillRecs(i, mvarSAPTxtRec)
				Case psiINDEX : psiCount = structFillRecs(i, mvarSAPPsiRec)
			End Select
		Next i
		hasOrderBytes = buildTree()
		readOrderBytes = hasOrderBytes
		If Not hasOrderBytes Then mvarOrderOpen = 0
		If mvarNeedsLogs Then mvarOrderLogs.Clear()
	End Function


	Private Sub Pause(ByRef TenthOfSec As Integer)
		Dim i As Integer

		On Error Resume Next
		For i = 1 To TenthOfSec
			Sleep(99)
		Next i
	End Sub


	Public Function getFieldValue(ByRef fieldName As String, ByRef targetString As String) As Boolean
		Dim i As Integer
		Dim t As String
		Dim tmp As String
		Dim res As String
		Dim result As Boolean
		result = False
		If mvarActField.mfLabel <> fieldName Then searchField(fieldName)
		If mvarActField.mfLabel = fieldName Then
			tmp = Trim(Mid(mvarSAPRecord(mvarActField.mfRecord), mvarActField.mfFirst, mvarActField.mfLength))
			If tmp <> "" Then
				result = CBool(Trim(CStr(convertTyp(res, mvarActField.mfLabel, mvarActField.mfTyp, tmp, "S"))))
				t = mvarActField.mfTyp
				If t = "S" Then t = "N"
				If t = "M" Then t = "N"
				If t = "N" Then
					result = CBool(stringForDB(CStr(convertTyp(targetString, "", t, res, "S"))))
				Else
					targetString = res
				End If
			Else
				targetString = ""
				result = True
			End If
		End If
		getFieldValue = result
	End Function
	Public Function getSapValue(ByRef field As MITAFIELD, Optional ByVal record As String = Nothing) As String
		If IsNothing(record) Then
			Return Trim(Mid(mvarSAPRecord(field.mfRecord), field.mfFirst, field.mfLength))
		Else
			Return Trim(Mid(record, field.mfFirst, field.mfLength))
		End If
	End Function
	Public Function getFieldContent(ByRef fieldName As String, ByRef targetString As String) As Boolean
		Dim i As Integer
		Dim t As String
		Dim tmp As String
		Dim res As String
		Dim result As Boolean
		result = False
		If mvarActField.mfLabel <> fieldName Then searchField(fieldName)
		If mvarActField.mfLabel = fieldName Then
			targetString = Trim(Mid(mvarSAPRecord(CInt(CStr(mvarActField.mfRecord))), mvarActField.mfFirst, mvarActField.mfLength))
			result = True
		End If
		getFieldContent = result
	End Function


	Public Function setFieldValue(ByRef fieldName As String, ByRef fieldContent As String) As Boolean
		Dim i As Integer
		Dim t As String
		Dim tmp As String
		setFieldValue = False
		If mvarActField.mfLabel <> fieldName Then searchField(fieldName)
		If mvarActField.mfLabel = fieldName Then
			tmp = "*****"
			Select Case mvarActField.mfTyp
				Case "N"
					tmp = Format(CShort(fieldContent), New String(CChar("0"), mvarActField.mfLength))
				Case "C"
					tmp = CStr(fieldContent)
				Case "S"
					tmp = Format(CInt(fieldContent) * 1000, New String(CChar("0"), mvarActField.mfLength))
				Case "M"
					tmp = Format(CInt(fieldContent) * 100, New String(CChar("0"), mvarActField.mfLength))
				Case "F"
					tmp = rad36Encode(CStr(fieldContent))
				Case Else
					tmp = tmp
			End Select
			If tmp <> "*****" Then
				Mid(mvarSAPRecord(mvarActField.mfRecord), mvarActField.mfFirst, mvarActField.mfLength) = tmp
				setFieldValue = True
			End If
		End If
	End Function


	Private Function stringForDB(ByRef buf1 As String) As String
		Dim buf2 As New System.Text.StringBuilder(buf1, 1000)
		buf2.Replace("'", " ")
		buf2.Replace("@", " ")
		buf2.Replace("´", " ")
		buf2.Replace("`", " ")
		stringForDB = buf2.ToString
	End Function


	Public Function getStatusValue(ByRef statLevel As Integer, ByRef statLabel As String, ByRef statContent As String) As Boolean
		Dim tmpStat As String
		Dim i As Integer
		Dim label As String
		Dim content As String
		getStatusValue = False
		Select Case statLevel
			Case 3
				tmpStat = mvarSAPStatRec(PAP(actPapIndex).statINDEX3)
			Case 7
				tmpStat = mvarSAPStatRec(MO(actMoINDEX).statINDEX7)
			Case 8
				tmpStat = mvarSAPStatRec(PS(actPsINDEX).statINDEX8)
		End Select
		tmpStat = Trim(Mid(tmpStat, mvarStatTableBegin))
		i = 1
		Do
			If i + mvarStatLabelLength + mvarStatContentLength + 1 > Len(tmpStat) Then Exit Function
			label = Trim(UCase(Mid(tmpStat, i, mvarStatLabelLength)))
			content = Trim(Mid(tmpStat, i + mvarStatLabelLength, mvarStatContentLength))
			If label = UCase(statLabel) Then
				statContent = content
				getStatusValue = True
				Exit Do
			End If
			i = i + mvarStatLabelLength
		Loop
	End Function



	Private Function convertTyp(ByRef target As String, ByRef FldNam As String, ByVal ConvTyp As String, ByRef KeyVal As String, ByRef AtxSap As String) As Boolean
		Dim i As Integer
		Dim buf, buffer As String
		Dim KeyNam(1) As String
		Dim FldNr As Integer
		Dim buf1() As String
		Dim KeyV As String
		Dim TmpDat As Date
		Dim KWert, FWert As String
		Dim result As Boolean
		convertTyp = True
		Select Case ConvTyp
			Case "N"
				buf = Trim(Str(CInt(KeyVal)))
			Case "C", "T"
				buf = KeyVal
			Case "S"
				If KeyVal <> "" Then
					If AtxSap = "S" Then
						buf = Trim(Str(CInt(KeyVal) / 1000))
					Else
						buf = Trim(Str(CInt(KeyVal) * 1000))
					End If
				End If
			Case "M"
				If KeyVal <> "" Then
					If AtxSap = "S" Then
						buf = Trim(Str(CInt(KeyVal) / 100))
					Else
						buf = Trim(Str(CInt(KeyVal) * 100))
					End If
				End If
			Case "U"
				If KeyVal <> "" Then
					If AtxSap = "S" Then
						buf = Trim(Str(CInt(KeyVal) / 100))
					Else
						buf = Trim(Str(CInt(KeyVal) * 100))
					End If
				End If
			Case "D"
				If AtxSap = "S" Then
					buf = Mid(KeyVal, 7, 2) & "/" & Mid(KeyVal, 5, 2) & "/" & Left(KeyVal, 4)
				Else
					TmpDat = CDate(KeyVal)
					buf = Format(TmpDat, "YYYYMMDD")
				End If
			Case "F"
				If AtxSap = "S" Then
					buf = rad36Decode(KeyVal)
				Else
					buf = rad36Encode(KeyVal)
				End If
			Case "O"
				i = InStr(KeyVal, "_")
				If i > 0 Then
					buf = Mid(KeyVal, i + 1)
				Else
					buf = " "
				End If
				i = InStr(buf, ".")
				If i > 0 Then buf = Left(buf, i - 1)
				buf = UCase(buf)
			Case "P"
				i = InStr(KeyVal, "_")
				If i > 0 Then
					buf = Left(KeyVal, i - 1)
				Else
					buf = KeyVal
				End If
				i = InStr(buf, ".")
				If i > 0 Then buf = Left(buf, i - 1)
				buf = UCase(buf)
			Case Else			 ' table
				If KeyVal > "" Then
					If InStr(ConvTyp, ".") = 0 Then
						buf = ""
						If IsNumeric(KeyVal) Then
							KeyVal = CStr(CInt(KeyVal))
						End If
						result = retreiveCustTableEntry(ConvTyp, KeyVal, AtxSap = "S", buf)
						If Not result Then
							If AtxSap = "S" Then
								result = retreiveCustTableEntry(ConvTyp, "DEFAULT", True, buf)
							End If
						End If
					Else
						i = InStr(ConvTyp, ".")
						FldNr = CInt(Mid(ConvTyp, i + 1))
						If FldNr > 0 Then FldNr = FldNr - 1
						ConvTyp = Left(ConvTyp, i - 1)
						result = retreiveCustTableEntry(ConvTyp, KeyVal, AtxSap = "S", buffer)
						buf1 = Split(buffer, ",")
						If UBound(buf1) >= FldNr Then
							buf = buf1(FldNr)
						Else
							buf = ""
						End If
					End If
					If buf = "" Then
						convertTyp = False
						eventProcess("Table-Entry '[" & ConvTyp & "] " & KeyVal & "' missing", (mitaEventCodes.errorUserIni), "convertTyp")
					End If
				Else
					buf = ""
				End If
		End Select
		If buf > "" Or ConvTyp = "X" Then
			target = Trim(buf)
		Else
			target = KeyVal
		End If
	End Function

	Private Function rad36Encode(ByRef FilNam As String) As String
		Dim DirNo, DatNo As Integer
		Dim ext As String
		Dim i As Integer
		Dim il, il1 As Integer
		Dim buffer As String

		i = InStr(FilNam, ".")
		If i > 0 Then
			ext = Mid(FilNam, i)
			FilNam = Left(FilNam, i - 1)
		End If
		i = InStr(FilNam, "\")
		DirNo = 0
		If i > 0 Then
			DirNo = CInt(Left(FilNam, i - 1))
		Else
			rad36Encode = ""
			Exit Function
		End If
		DatNo = CInt(Mid(FilNam, i + 1))
		rad36Encode = toRAD36(DirNo, 3) & toRAD36(DatNo, 5) & ext
	End Function

	Private Function toRAD36(ByRef Zahl As Integer, ByRef cnt As Integer) As String
		Dim buffer As String
		Dim i As Integer
		Dim il, il1 As Integer

		il = 36
		For i = 1 To cnt - 1
			il = il * 36
		Next i
		buffer = ""
		For i = 0 To cnt
			il1 = Zahl \ il
			If il1 < 10 Then
				buffer = buffer & Trim(Str(il1))
			ElseIf il1 < 36 Then
				buffer = buffer & Chr(((il1 - 10) Mod 26) + 65)
			Else
				buffer = buffer & "0"
			End If
			Zahl = Zahl - il1 * il
			il = CInt(il / 36)
		Next i
		toRAD36 = Right(New String(CChar("0"), cnt) & buffer, cnt)
	End Function

	Private Function rad36Decode(ByRef FilNam As String) As String
		Dim DirNo, DatNo As Integer
		Dim i As Integer
		Dim ext As String

		On Error Resume Next
		If Len(FilNam) < 4 Then
			rad36Decode = ""
			Exit Function
		End If
		FilNam = UCase(FilNam)
		i = InStr(FilNam, ".")
		If i > 0 Then
			ext = Mid(FilNam, i)
			FilNam = Left(FilNam, i - 1)
		Else
			ext = ".ADV"
		End If
		DirNo = 0
		For i = 1 To 3
			If Mid(FilNam, i, 1) >= "0" And Mid(FilNam, i, 1) <= "9" Then
				DirNo = DirNo * 36 + Asc(Mid(FilNam, i, 1)) - 48
			ElseIf Mid(FilNam, i, 1) >= "A" And Mid(FilNam, i, 1) <= "Z" Then
				DirNo = DirNo * 36 + Asc(Mid(FilNam, i, 1)) - 55
			Else
				DirNo = DirNo * 36
			End If
		Next i
		DatNo = 0
		For i = 4 To Len(FilNam)
			If Mid(FilNam, i, 1) >= "0" And Mid(FilNam, i, 1) <= "9" Then
				DatNo = DatNo * 36 + Asc(Mid(FilNam, i, 1)) - 48
			ElseIf Mid(FilNam, i, 1) >= "A" And Mid(FilNam, i, 1) <= "Z" Then
				DatNo = DatNo * 36 + Asc(Mid(FilNam, i, 1)) - 55
			Else
				DatNo = DatNo * 36
			End If
		Next i

		rad36Decode = Trim(Str(DirNo)) & "\" & Format(DatNo, "####0000") & ext
	End Function


	Private Function orderNextAllInsertion() As Boolean
		Dim cnt As Integer
		Dim i As Integer
		Dim tmp As INSERTION
		orderNextAllInsertion = False
		If Not checkOpen() Then Exit Function
		If actPsINDEX = psCount Then
			Exit Function
		End If
		actPsINDEX = actPsINDEX + 1
		tmp = PS(actPsINDEX)
		actPlzINDEX = tmp.plzINDEX
		actPlzaINDEX = tmp.plzaINDEX
		actPsiINDEX = tmp.psiINDEX8
		orderNextAllInsertion = True
	End Function
	Public Function itemTestLastInsertion() As Boolean
		Return actPsINDEX = psCount
	End Function
	Public Function itemTestLastPub() As Boolean
		Return actPapIndex = papCount
	End Function
	Public Function itemTestLastCombo() As Boolean
		Return comboIndex = comboCount
	End Function
	Public Function itemTestLastMotiv() As Boolean
		Return actMoINDEX = moCount
	End Function
	Public Function orderNextInsertion() As Boolean
		Dim i As Integer
		Dim tmp As INSERTION
		Dim index As Integer = PAP(actPapIndex).psIndex + 1
		orderNextInsertion = False
		If index > PAP(actPapIndex).psCount Then Exit Function
		mvarRequeryReturn = mitaEventReturnCodes.returnOk
		mvarRequeryCount = 0
		If Not checkOpen() Then Exit Function
		PAP(actPapIndex).psIndex = index
		actPsINDEX = PAP(actPapIndex).psIndexes(PAP(actPapIndex).psIndex)
		mvarSAPRecord(psINDEX) = mvarSAPPsRec(actPsINDEX)
		tmp = PS(actPsINDEX)
		actPlzINDEX = tmp.plzINDEX
		actPlzaINDEX = tmp.plzaINDEX
		actBlzINDEX = tmp.blzINDEX
		If actBlzINDEX <> -1 Then mvarSAPRecord(blzINDEX) = mvarSAPBlzRec(actBlzINDEX)
		If actPlzINDEX <> -1 Then mvarSAPRecord(plzINDEX) = mvarSAPPlzRec(actPlzINDEX)
		If actPlzaINDEX <> -1 Then mvarSAPRecord(plzaINDEX) = mvarSAPPlzaRec(actPlzaINDEX)
		actPsiINDEX = tmp.psiINDEX8
		mvarInsertionCount = mvarInsertionCount + 1
		orderNextInsertion = True
	End Function

	Private Function orderNextAllMotiv() As Boolean
		Dim cnt As Integer
		orderNextAllMotiv = False
		If Not checkOpen() Then Exit Function
		If actMoINDEX = moCount Then
			Exit Function
		End If
		actMoINDEX = actMoINDEX + 1
		If MO(actMoINDEX).moNr = 0 Then
			Exit Function
		End If
		mvarMotivNo = getSapValue(mvarFldOrderMotiv)
		'mvarSapNo = mvarOrderNo & mvarMotivNo
		orderResetCombo()
		orderNextAllMotiv = True
	End Function
	Public Function orderNextMotiv() As Boolean
		'mvarSapNo = ""
		orderNextMotiv = False
		mvarMotivNo = CStr(0)
		mvarTpAdNo = ""
		mvarPaper = ""
		mvarTpAdVer = ""
		orderNextMotiv = False
		If Not checkOpen() Then Exit Function
		mvarStartMotivTime = VB.Timer()
		Do
			actMoINDEX = actMoINDEX + 1
			If actMoINDEX > moCount Then
				Exit Function
			End If
			If MO(actMoINDEX).used Then
				Exit Do
			End If
		Loop
		mvarSAPRecord(moINDEX) = mvarSAPMoRec(actMoINDEX)
		mvarMotivNo = getSapValue(mvarFldOrderMotiv)
		'mvarSapNo = mvarOrderNo & mvarMotivNo
		mvarEinNo = getSapValue(mvarFldProduct)
		mvarPaper = Trim(mvarEinNo)
		combo = infoGetcombis(mvarPaper)
		comboCount = UBound(combo)
		If comboCount > 0 Then
			mvarCombo = mvarPaper
		Else
			mvarCombo = ""
		End If
		comboIndex = -1
		orderResetPub()
		moCountUsed = moCountUsed + 1
		orderNextMotiv = True
	End Function

	Public Function orderResetCombo() As Boolean
		Dim cnt As Integer
		orderResetCombo = False
		If Not checkOpen() Then Exit Function
		comboIndex = -1
		orderResetCombo = True
	End Function

	Public Function orderResetMotiv() As Boolean
		Dim cnt As Integer
		orderResetMotiv = False
		If Not checkOpen() Then Exit Function
		If moCount = -1 Then
			Exit Function
		End If
		actPakINDEX = -1
		actPapIndex = -1
		actBpzINDEX = -1
		actIszINDEX = -1
		actMoINDEX = -1
		actBlzINDEX = -1
		actPsINDEX = -1
		actPlzINDEX = -1
		actPlzaINDEX = -1
		actStatINDEX = -1
		actTxtINDEX = -1
		actPsiINDEX = -1
		actErrINDEX = -1
		mvarMotivNo = "000000"
		mvarTpAllAd = ""
		mvarTpNumAd = 0
		moCountUsed = 0
		orderResetMotiv = True
	End Function

	Public Function orderResetInsertion() As Boolean
		Dim cnt As Integer
		orderResetInsertion = False
		If Not checkOpen() Then Exit Function
		If psCount = -1 Then
			Exit Function
		End If
		PAP(actPapIndex).psIndex = -1
		mvarInsertionCount = 1
		mvarEinNo = "000000"
		orderResetInsertion = True
	End Function

	Public Function orderResetPub() As Boolean
		Dim cnt As Integer
		orderResetPub = False
		If Not checkOpen() Then Exit Function
		If papCount = -1 Then
			Exit Function
		End If
		MO(actMoINDEX).papIndex = -1
		actPapIndex = -1
		mvarPosNo = "000"
	End Function

	Public Function orderNextCombo() As Boolean
		Dim cnt As Integer
		orderNextCombo = False
		If Not checkOpen() Then Exit Function
		If comboIndex = comboCount Then Exit Function
		comboIndex = comboIndex + 1
		mvarPaper = combo(comboIndex)
		orderResetInsertion()
		orderNextCombo = True
	End Function

	Public Function orderNextPub() As Boolean
		orderNextPub = False
		If Not checkOpen() Then Exit Function
		MO(actMoINDEX).papIndex = MO(actMoINDEX).papIndex + 1
		If MO(actMoINDEX).papIndex > MO(actMoINDEX).papCount Then Exit Function
		actPapIndex = MO(actMoINDEX).papIndexes(MO(actMoINDEX).papIndex)
		actIszINDEX = actPapIndex
		actBpzINDEX = actPapIndex
		mvarSAPRecord(papIndex) = mvarSAPPapRec(actPapIndex)
		mvarSAPRecord(iszINDEX) = mvarSAPIszRec(actPapIndex)
		mvarSAPRecord(bpzINDEX) = mvarSAPBpzRec(actPapIndex)
		If mvarCombo = "" Then
			orderResetInsertion()
		Else
			orderResetCombo()
		End If
		mvarPosNo = getSapValue(mvarFldPAPPos, mvarSAPRecord(papIndex))
		mvarRefAVM = getSapValue(mvarRefPAPAVM, mvarSAPRecord(papIndex))
		mvarRefPos = getSapValue(mvarRefPAPPos, mvarSAPRecord(papIndex))
		orderNextPub = True
	End Function

	Private Function structFillRecs(ByRef mvarnt As Integer, ByRef Arr() As String) As Integer
		Dim l As Integer
		Dim i As Integer
		l = mvarSapRecords(mvarnt).mtCount - 1
		If l < 0 Then
			ReDim Arr(0)
		Else
			ReDim Arr(l)
			For i = 0 To l
				Arr(i) = mvarSapRecords(mvarnt).mtRecords(i).sbContent
			Next i
		End If
		structFillRecs = l
	End Function

	Public Sub New()
		MyBase.New()
		ReDim mvarSAPRecLen(1)
		ReDim mvarSAPRecName(1)
		ReDim mvarSAPRecord(1)
		mvarIniLoaded = False
		mvarOrderOpen = 0
		hasOrderBytes = False
		mvarLogClassMap = 0
		Dim tmp As String = System.Windows.Forms.Application.ExecutablePath
		Dim i As Integer = InStrRev(tmp, "\")
		mvarParentPath = Left$(tmp, i)
		mvarBreakFlag = False
	End Sub

	Private Function dbCloseOdbc() As Boolean
		Return mitaConnect.connectionCloseOdbc()
	End Function
	Private Function dbOpenOdbc() As Boolean
		Return mitaConnect.connectionOpenOdbc()
	End Function
	Private Function dbTest(ByRef dbConnectString As String) As String
		dbTest = mitaConnect.connectionTest(dbConnectString)
	End Function


	Protected Overrides Sub Finalize()
		dbCloseOdbc()
		mitaMessage.Close()
		MyBase.Finalize()
	End Sub


	'Public Function namesSetUser(ByRef id As String) As Boolean
	'	mitaData.createUser = stringForDB(Trim(id))
	'	namesSetUser = (mitaData.createUser <> "")
	'	If mitaData.createUser = "" Then
	'		eventProcess("Invalid data:" & id, (mitaEventCodes.errorInvalidInput), "namesSetUser")
	'		Exit Function
	'	End If
	'End Function

	'Public Function namesSetApplication(ByRef id As String) As Boolean
	'	mitaData.mitaApplication = stringForDB(Trim(id))
	'	namesSetApplication = (mitaData.mitaApplication <> "")
	'	If mitaData.mitaApplication = "" Then
	'		eventProcess("Invalid data:" & id, (mitaEventCodes.errorInvalidInput), "namesSetApplication")
	'		Exit Function
	'	End If
	'End Function

	'Public Function namesSetSapPool(ByRef pool As String) As Boolean
	'	If pool <> "" Then mitaData.sapPool = Trim(pool)
	'	Return True
	'End Function

	'Public Function namesSetProcessId(ByRef id As Integer) As Boolean
	'	mvarProcessID = id
	'	Return True
	'End Function
	'Public Function namesSetCaption(ByRef caption As String) As Boolean
	'	mitaData.caption = caption
	'	Return True
	'End Function
	'Public Function getCustomerQuery(ByRef SrcSQL As String, ByRef targetTemplate As String, ByRef targetSQL As String) As Boolean
	Public Function getCustomerQuery(ByRef targetSQL As SQLLOG) As Boolean
		Dim buffer As String
		Dim NewSQL As New System.text.StringBuilder(1000)
		Dim Sql As String
		Dim tmp As String
		Dim iPos, i As Integer
		Dim Index As Integer
		Dim ok As Boolean
		Dim searchStart As Integer
		getCustomerQuery = False
		Index = searchQuery(targetSQL.sName)
		If Index = -1 Then
			eventProcess("Query Entry " & targetSQL.sName & " missing", (mitaEventCodes.errorUserIni), "getCustomerQuery")
			Exit Function
		End If
		targetSQL.sPaper = mvarPaper
		targetSQL.sTemplate = mvarCustomQueries(Index)
		targetSQL.sError = ""
		searchStart = 1
		Sql = mvarCustomQueries(Index)
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
				NewSQL.Append("#" & tmp & "#")
			ElseIf Left(tmp, 4) = "MITA" Then
				buffer = Mid(tmp, 5)
				If buffer = "ORDERNO" Then
					NewSQL.Append(mvarOrderNo)
				ElseIf buffer = "VNO" Then
					NewSQL.Append(CStr(CInt(mvarVNO)))
				ElseIf buffer = "MOTIVNO" Then
					NewSQL.Append(CStr(CInt(mvarMotivNo)))
				ElseIf buffer = "SAPNO" Then
					NewSQL.Append(mvarOrderNo & mvarMotivNo)
				ElseIf buffer = "INSCOUNT" Then
					NewSQL.Append(CStr(CInt(PAP(actPapIndex).psCount + 1)))
				ElseIf buffer = "INSINDEX" Then
					NewSQL.Append(CStr(CInt(mvarInsertionCount)))
				ElseIf buffer = "POSNO" Then
					NewSQL.Append(CStr(CInt(mvarPosNo)))
				ElseIf buffer = "EINNO" Then
					NewSQL.Append(CStr(CInt(mvarEinNo)))
				ElseIf buffer = "PAPER" Then
					NewSQL.Append(mvarPaper)
				ElseIf buffer = "COMBONO" Then
					NewSQL.Append(CStr(CInt(comboIndex)))
				ElseIf buffer = "ID" Then
					NewSQL.Append(CStr(mitaData.createID))
				ElseIf buffer = "USER" Then
					NewSQL.Append(mitaData.createUser)
				ElseIf buffer = "CLIENTNO" Then
					NewSQL.Append(mitaSystem.sapSystemSAPCLIENT)
				ElseIf buffer = "POOL" Then
					NewSQL.Append(mitaSystem.sapPool)
				ElseIf buffer = "REFAVM" Then
					NewSQL.Append(mvarRefAVM)
				ElseIf buffer = "REFPOS" Then
					NewSQL.Append(CStr(CInt(mvarRefPos)))
				End If
			ElseIf Left(tmp, 4) = "STAT" Then
				i = CInt(Mid(tmp, 5, 1))
				ok = getStatusValue(i, Mid(tmp, 6), buffer)
				If Not ok Then
					RaiseEvent sqlError("STAT Field not found: " & tmp, "", "User Mistake")
					RaiseEvent endApplication(True)
					Exit Function
				End If
				NewSQL.Append(buffer)
			Else
				ok = getFieldValue(tmp, buffer)
				If Not ok Then
					RaiseEvent sqlError("Customer Field not found: " & tmp, "", "User Mistake")
					RaiseEvent endApplication(True)
					Exit Function
				End If
				NewSQL.Append(buffer)
			End If
		Loop
		targetSQL.sResult = Replace(NewSQL.ToString, "§", mitaSystem.sapSystemId)
		getCustomerQuery = True
	End Function

	Public Function iniReset() As Boolean
		mvarIniLoaded = False
		Return True
	End Function
	Private Sub eventSAPERRORS(ByRef parameter As String)
		If mvarSapError.getErrorCount > -1 Then
			Dim blobs As Long
			Dim sendDirect As Boolean
			mvarStartErrorTime = VB.Timer()
			If parameter <> "" Then
				mvarStartErrorTime = VB.Timer()
			End If
			sendDirect = False
			RaiseEvent sapErrorForSend(sendDirect)
			If sendDirect Then
				mvarSapError.sapErrorsSend()
			Else
				mvarErrorBytes = mvarSapError.errorRecords
				mvarErrorByteCount = mvarErrorBytes.Length
				Dim result As Boolean = errorWriteDB(blobs)
			End If
			If parameter <> "" Then
				Dim Elg As String = eventFillData(parameter, "", Nothing, "")
				RaiseEvent logContent(Elg, "L")
			End If
		End If
		mvarSapError.sapResetErrors()
	End Sub

	Private Sub eventSAPERR(ByVal parameter As String)
		Dim code As Integer
		Dim msg As String
		Dim struct As String
		Dim x() As String
		If mvarSapError Is Nothing Then
			eventRaise("MitaError not set", mitaEventCodes.errorProgrammer, "eventSAPERR")
			RaiseEvent endApplication(True)
		End If
		x = Split(parameter, "|")
		msg = Trim(x(0))
		code = CShort(x(1))
		mvarSapError.sapAddError(msg, code, mvarOrderNo, mvarPosNo, mvarEinNo, mvarMotivNo, mvarVNO, mitaData.createUser)
	End Sub

	Public Function namesSetSapError(ByVal sapError As Object) As Boolean
		mvarSapError = CType(sapError, PscMitaError.CMitaError)
		mvarSapError.dataSet = mitaData
		mvarSapError.sharedSet = mitaShared
		mvarSapError.connectSet = mitaConnect
		Return True
	End Function

	'Public Function namesSetRunType(ByVal rType As String) As Boolean
	'	mitaData.runType = rType
	'	Return True
	'End Function

	'Public Sub orderStartPoll()
	'	AddHandler aTimer.Elapsed, AddressOf OnTimedEvent
	'	aTimer.Interval = 10
	'	aTimer.Enabled = True
	'End Sub

	'Public Sub orderEndPoll()
	'	aTimer.Close()
	'	aTimer.Dispose()
	'End Sub
	'Private Sub OnTimedEvent(ByVal source As Object, ByVal e As System.timers.ElapsedEventArgs)
	'	CType(source, System.Timers.Timer).Stop()
	'	If mvarOrderOpen = 0 Then
	'		CType(source, System.Timers.Timer).Enabled = True
	'		If orderReadNextDB() Then
	'			RaiseEvent orderArrived()
	'		End If
	'	End If
	'	CType(source, System.Timers.Timer).Start()
	'End Sub
	Private Function retreiveCustTableEntry(ByVal tableName As String, ByVal KeyVal As String, ByVal isFromSap As Boolean, ByRef buf As String) As Boolean
		Dim i As Integer
		For i = 0 To mvarCustTableCount
			If mvarCustTablesArray(i).name = tableName Then Exit For
		Next
		If i > mvarCustTableCount Then
			mvarCustTableCount = mvarCustTableCount + 1
			ReDim Preserve mvarCustTablesArray(mvarCustTableCount)
			mvarCustTablesArray(mvarCustTableCount) = readDBCustTableEntries(tableName)
			mvarCustTablesArray(mvarCustTableCount).name = tableName
		End If
		buf = searchTableValue(mvarCustTablesArray(i), KeyVal, isFromSap)
		Return (Not (IsNothing(buf)))
	End Function
	Private Function searchTableValue(ByVal entry As custTableArray, ByVal srch As String, ByVal isLeft As Boolean) As String
		Dim von As Integer
		Dim bis As Integer
		Dim test As Integer
		Dim lftArr() As String
		Dim rthArr() As String
		von = LBound(entry.tableLeftLeft)
		bis = UBound(entry.tableLeftLeft)
		If isLeft Then
			lftArr = entry.tableLeftLeft
			rthArr = entry.tableLeftRight
		Else
			rthArr = entry.tableRightLeft
			lftArr = entry.tableRightRight
		End If
		Do
			If lftArr(von) = srch Then
				Return rthArr(von)
			End If
			If lftArr(bis) = srch Then
				Return rthArr(bis)
			End If
			If bis - von <= 1 Then
				Exit Function
			End If
			test = CInt((bis - von) / 2 + von)
			If lftArr(test) = srch Then
				Return rthArr(test)
			End If
			If lftArr(test) > srch Then
				bis = test
			Else
				von = test
			End If
		Loop
	End Function
	Public Function orderReadSap() As Boolean
		Dim add As String
		'Dim struct() As sapStructure
		Dim rfcrc As Integer
		Dim cStruct() As STRUCTSTRUCT
		Dim cCount As Integer
		If mitaData.isMiniSAP Then
			If mitaSystem.sapSystemVERSIONID = 1 Then
				add = "Z"
			Else
				add = "ZZ"
			End If
		Else
			add = ""
		End If
		If IsNothing(mvarSapServer) Then
			mvarSapServer = New pscSapServer.CSapServer
			If Not readDBStructures(cStruct, cCount, rfcClass.rfcOrder) Then Return Nothing
			mvarOrderFrame = mvarSapServer.createFrame(add, rfcClass.rfcOrder, cStruct, cCount)
			'struct = mvarOrderFrame.cfDataTables
			'struct(0).csStructure = add & x(1)
			'struct(0).csCount = UBound(mvarSAPErrRec)
			'struct(0).csData = mvarSAPErrRec
			'mvarOrderFrame.cfDataTables = struct
			mvarOrderFrame.cfSystem = mitaSystem.sapSystemSAPSYSTEM
			mvarOrderFrame.cfHost = mitaData.createHost
			mvarOrderFrame.cfProgram = mitaData.mitaApplication
			mvarOrderFrame.cfGateway = mitaSystem.sapSystemSAPGATEWAY
			mvarOrderFrame.cfService = mitaSystem.sapSystemSAPSERVICE
			mvarOrderFrame.cfUser = mitaSystem.sapSystemSAPUSER
			mvarOrderFrame.cfPassword = mitaSystem.sapSystemSAPOWNER
			mvarOrderFrame.cfServer = mitaSystem.sapSystemSAPSERVER
			mvarOrderFrame.cfMandant = mitaSystem.sapSystemSAPCLIENT
			mvarOrderFrame.cfLanguage = "EN"
			mvarOrderFrame.cfTrace = mitaData.doTrace
			mvarOrderFrame.cfFunction = add & cStruct(0).sFunction
			mvarOrderFrame.cfTyp = rfcClass.rfcOrder
			If mvarSapServer.sapInitInput(mvarOrderFrame) = 0 Then
				eventRaise(mvarSapServer.getLastError(), mitaEventCodes.errorSAPConnection)
				Return False
			End If
		End If
		If Not mvarSapServer.orderGetNext() Then
			Return False
		End If
		Return True
	End Function
	Public Function orderGetErrorCode() As mitaErrorCodes
		Return mvarOrderError
	End Function
	Public Function orderTransfer() As Boolean
		Return False
	End Function
	Public Function optionsSetSapTrace(ByVal trace As Char) As Boolean
		Select Case trace
			Case "0"c, "1"c
				mitaData.doTrace = CInt(Val(trace))
			Case "D"c
				mitaData.doTrace = Asc(trace)
			Case Else
				mitaData.doTrace = 0
		End Select
		Return True
	End Function
	Property messageSet() As pscMitaMsg.CMitaMsg
		Get
			Return mitaMessage
		End Get
		Set(ByVal Value As pscMitaMsg.CMitaMsg)
			mitaMessage = Value
		End Set
	End Property
	Property dataSet() As pscMitaData.CMitaData
		Get
			Return mitaData
		End Get
		Set(ByVal Value As pscMitaData.CMitaData)
			mitaData = Value
			If IsNothing(mitaData.createHost) Then mitaData.createHost = System.Net.Dns.GetHostName
			If Not IsNothing(mitaConnect) Then
				Dim errorMessage As String = dbTest(mitaConnect.dbConnectString)
				If errorMessage = "" Then
					dbOpenOdbc()
					dbCloseOdbc()
					mitaData.connectionOK = True
				End If
			End If
		End Set
	End Property
	Property systemSet() As pscMitaSapSystem.CMitaSapSystem
		Get
			Return mitaSystem
		End Get
		Set(ByVal Value As pscMitaSapSystem.CMitaSapSystem)
			mitaSystem = Value
		End Set
	End Property
	Property sharedSet() As pscMitaShared.CMitaShared
		Get
			Return mitaShared
		End Get
		Set(ByVal Value As pscMitaShared.CMitaShared)
			mitaShared = Value
		End Set
	End Property
	Property connectSet() As pscMitaConnect.CMitaConnect
		Get
			Return mitaConnect
		End Get
		Set(ByVal Value As pscMitaConnect.CMitaConnect)
			mitaConnect = Value
			If Not IsNothing(mitaData) Then
				Dim errorMessage As String = dbTest(Value.dbConnectString)
				If errorMessage = "" Then
					dbOpenOdbc()
					dbCloseOdbc()
					mitaData.connectionOK = True
				End If
			End If
		End Set
	End Property
	Private Function errorReadDB(ByRef orderID As Integer) As Boolean
		Dim query As String
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		resetAll()
		mvarOrderID = orderID
		If Not mvarIniLoaded Then iniRead()
		If Not checkInputs() Then Exit Function
		mvarStartOrderTime = VB.Timer()
		eventProcess("", mitaEventCodes.programOrderRead, "orderReadDB")
		errorReadDB = False
		query = sqlTrans(mvarFind)
		idbc.CommandText = query
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		If reader.Read Then
			mvarErrorByteCount = reader.GetInt32(0)
			ReDim mvarErrorBytes(mvarErrorByteCount - 1)
			reader.GetBytes(1, 0, mvarErrorBytes, 0, mvarErrorByteCount)
			mvarOrderNo = reader.GetString(2)
			mvarVNO = CStr(reader.GetInt32(3))
			reader.Close()
			mvarErrorOpen = 1
			errorReadDB = True
			mvarErrorOpen = -1
		Else
			reader.Close()
		End If
		mitaConnect.odbc_connection.Close()
	End Function
	Public Property tools() As Boolean
		Get
			Return mvarIsTools
		End Get
		Set(ByVal Value As Boolean)
			mvarIsTools = Value
		End Set
	End Property
End Class

