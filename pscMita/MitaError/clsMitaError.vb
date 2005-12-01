Option Strict On
Option Explicit On 
Imports Microsoft.VisualBasic
Imports pscSapClient.CSapClient
Imports pscMitaDef.CMitaDef

'<ComClass(MitaError.ClassId, MitaError.InterfaceId, MitaError.EventsId)> 
Public Class CMitaError
	'#Region "COM GUIDs"
	'	' These  GUIDs provide the COM identity for this class 
	'	' and its COM interfaces. If you change them, existing 
	'	' clients will no longer be able to access the class.
	'	Public Const ClassId As String = "EE1BE61F-8332-45c6-A081-2A90550248ED"
	'	Public Const InterfaceId As String = "FBC70D51-FB92-4b56-BC66-35D76C964E74"
	'	Public Const EventsId As String = "07A68988-D779-4e6a-B966-6F9036D0502D"
	'#End Region
	Private mvarSAP2DB() As MITAFIELD
	Private mvarSapRecords() As MITATABLE

	Private mvarSAPErrRec() As String
	Private mvarErrorCountSAP As Integer
	Private mvarErr2Sap() As MITAFIELD
	Private mvarErrDef() As MITAFIELD
	Private mvarErrMsgField As MITAFIELD
	Private mvarIniLoaded As Boolean
	Private mvarFunctionName As String
	Private mvarStructureName As String
	'Private mvarSapSystem As String
	Private mvarErrorFunction As String
	Private mvarErrorStructure As String

	Public Event sapError(ByRef message As String, ByVal code As Integer)

	Public Sub New()
		sapResetErrors()
	End Sub

	Public Sub sapAddError(ByRef description As String, ByVal code As Integer, _
	ByRef avmNr As String, ByVal posNr As String, ByVal einNr As String, _
	ByVal motiv As String, ByVal vno As String, ByRef user As String)
		Dim i, i1 As Integer
		Dim buf As String
		Dim x() As String
		Dim typ As String
		Dim content As String
		If Val(motiv) = 0 Then motiv = Nothing
		If Not mvarIniLoaded Then iniRead()
		mvarErrorCountSAP = mvarErrorCountSAP + 1
		ReDim Preserve mvarSAPErrRec(mvarErrorCountSAP)
		mvarSAPErrRec(mvarErrorCountSAP) = Space(mvarSapRecords(errINDEX).mtLength)
		For i = 0 To UBound(mvarSAP2DB)
			buf = ""
			If mvarSAP2DB(i).mfLabel = "ERRDATE" Then
				buf = Format(Now, "yyyyMMdd")
			ElseIf mvarSAP2DB(i).mfLabel = "ERRTIME" Then
				buf = Format(Now, "HHmmss")
			ElseIf mvarSAP2DB(i).mfLabel = "ERRTYP" Then
				buf = Format(code, New String(CChar("0"), mvarSAP2DB(i).mfLength))
			ElseIf mvarSAP2DB(i).mfLabel = "ERRMSG" Then
				buf = description
			ElseIf mvarSAP2DB(i).mfLabel = "ERRAVMNR" Then
				If Not IsNothing(avmNr) Then buf = avmNr
			ElseIf mvarSAP2DB(i).mfLabel = "ERRPOSNR" Then
				If Not IsNothing(posNr) Then buf = posNr
			ElseIf mvarSAP2DB(i).mfLabel = "ERREINNR" Then
				If Not IsNothing(einNr) Then buf = einNr
			ElseIf mvarSAP2DB(i).mfLabel = "ERRMOTIV" Then
				If Not IsNothing(motiv) Then buf = motiv
			ElseIf mvarSAP2DB(i).mfLabel = "ERRVNO" Then
				If Not IsNothing(vno) Then buf = vno
			ElseIf mvarSAP2DB(i).mfLabel = "ERRUSERID" Then
				If Not IsNothing(user) Then buf = user
			End If
			If mvarSAP2DB(i).mfTyp = "N" Then
				Mid(mvarSAPErrRec(mvarErrorCountSAP), mvarSAP2DB(i).mfFirst, mvarSAP2DB(i).mfLength) = Right(New String(CChar("0"), mvarSAP2DB(i).mfLength) & buf, mvarSAP2DB(i).mfLength)
			Else
				Mid(mvarSAPErrRec(mvarErrorCountSAP), mvarSAP2DB(i).mfFirst, mvarSAP2DB(i).mfLength) = Left(buf & Space(mvarSAP2DB(i).mfLength), mvarSAP2DB(i).mfLength)
			End If
		Next i
	End Sub

	Public Function sapErrorsSend() As Boolean
		Dim client As New pscSapClient.CSapClient
		Dim fram As sapFrame
		Dim struct() As sapStructure
		Dim cStruct() As STRUCTSTRUCT
		Dim cCount As Integer
		If Not readDBStructures(cStruct, cCount, rfcClass.rfcError) Then Return Nothing
		fram = client.createFrame(rfcClass.rfcError, cStruct, cCount)
		fram.cfFunction = cStruct(cCount).sFunction
		struct = fram.cfDataTables
		struct(0).csStructure = cStruct(cCount).sData
		struct(0).csCount = UBound(mvarSAPErrRec)
		struct(0).csData = mvarSAPErrRec
		fram.cfCountTables = 0
		fram.cfDataTables = struct
		fram.cfSystem = mitaSystem.sapSystemSAPSYSTEM
		fram.cfGateway = mitaSystem.sapSystemSAPGATEWAY
		fram.cfService = mitaSystem.sapSystemSAPSERVICE
		fram.cfUser = mitaSystem.sapSystemSAPUSER
		fram.cfPassword = mitaSystem.sapSystemSAPOWNER
		fram.cfServer = mitaSystem.sapSystemSAPSERVER
		fram.cfMandant = mitaSystem.sapSystemSAPCLIENT
		fram.cfLanguage = "EN"
		fram.cfTrace = mitaData.doTrace
		fram.cfTyp = rfcClass.rfcError
		If Not client.clientSapSend(fram) Then
			RaiseEvent sapError(client.getLastError(), mitaEventCodes.errorSAPConnection)
			Return False
		End If
		Return True
	End Function


	Private Sub iniRead()
		Dim tmpRec As MITAFIELD
		Dim result As Boolean
		Dim cFields() As FIELDSTRUCT
		Dim cCount As Integer
		Dim i As Integer
		Dim j As Integer
		Dim k As Integer
		Dim a As String
		Dim structureCount As Integer
		Dim structures() As STRUCTSTRUCT
		ReDim mvarErr2Sap(1)
		ReDim mvarErrDef(1)
		mvarIniLoaded = False
		structureCount = -1
		result = readDBStructures(structures, structureCount, rfcClass.rfcError)
		If Not result Or structureCount < 0 Then Exit Sub
		ReDim mvarSapRecords(structureCount)
		For i = 0 To structureCount
			mvarSapRecords(i).mtName = structures(i).sName
			mvarSapRecords(i).mtLength = structures(i).sLength
			mvarSapRecords(i).mtExt = "." & Mid(mvarSapRecords(i).mtName, StrRecord)
		Next i
		cCount = -1
		result = readDBCustFields(cFields, cCount, rfcClass.rfcError)
		If Not result Or cCount < 0 Then Exit Sub
		ReDim mvarSAP2DB(cCount)
		For i = 0 To cCount
			tmpRec.mfFirst = cFields(i).fFirst
			tmpRec.mfLabel = cFields(i).fName
			tmpRec.mfLength = cFields(i).fLength
			tmpRec.mfTyp = cFields(i).fType
			For j = 0 To structureCount
				If cFields(i).fStructure = mvarSapRecords(j).mtName Then
					tmpRec.mfRecord = j
					Exit For
				End If
			Next j
			If i = 0 Then
				mvarSAP2DB(i) = tmpRec
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
		mvarIniLoaded = True
	End Sub
	Public Function getErrorCount() As Integer
		Return mvarErrorCountSAP
	End Function

	'Public Function namesSetSapSystemName(ByRef sapSystemName As String) As Boolean
	'	Dim result As Boolean
	'	mitaData.sapSystemNAME = sapSystemName
	'	result = readDBSapSystemFromName(mitaData.sapSystemNAME)
	'	Return result
	'End Function

	'Public Function namesSetSapVersion(ByRef sapVersion As Integer) As Boolean
	'	mitaData.sapSystemVERSIONID = sapVersion
	'	Return True
	'End Function

	Private Function dbCloseOdbc() As Boolean
		Return mitaConnect.connectionCloseOdbc()
	End Function

	Private Function dbOpenOdbc() As Boolean
		Return mitaConnect.connectionOpenOdbc()
	End Function

	Public Sub sapResetErrors()
		mvarErrorCountSAP = -1
		ReDim mvarSAPErrRec(0)
	End Sub

	Protected Overrides Sub Finalize()
		dbCloseOdbc()
		MyBase.Finalize()
	End Sub
	'Public Function namesSetConnectString(ByVal connectString As String) As String
	'	Dim errorMessage As String = dbTest(connectString)
	'	If errorMessage = "" Then
	'		mitaConnect.dbConnectString = connectString
	'		dbOpenOdbc()
	'		dbCloseOdbc()
	'		Return ""
	'	End If
	'	Return errorMessage
	'End Function
	Private Function dbTest(ByRef dbConnectString As String) As String
		dbTest = mitaConnect.connectionTest(dbConnectString)
	End Function
	'Public Function optionsSetSapTrace(ByVal trace As Char) As Boolean
	'	Select Case trace
	'		Case "0"c, "1"c
	'			mitaData.doTrace = CInt(Val(trace))
	'		Case "D"c
	'			mitaData.doTrace = Asc(trace)
	'		Case Else
	'			mitaData.doTrace = 0
	'	End Select
	'	Return True
	'End Function
	Public Property systemSet() As pscMitaSapSystem.CMitaSapSystem
		Get
			Return mitaSystem
		End Get
		Set(ByVal Value As pscMitaSapSystem.CMitaSapSystem)
			mitaSystem = Value
		End Set
	End Property

	Public Property dataSet() As pscMitaData.CMitaData
		Get
			Return mitaData
		End Get
		Set(ByVal Value As pscMitaData.CMitaData)
			mitaData = Value
		End Set
	End Property
	Public Property sharedSet() As pscMitaShared.CMitaShared
		Get
			Return mitaShared
		End Get
		Set(ByVal Value As pscMitaShared.CMitaShared)
			mitaShared = Value
		End Set
	End Property
	Public Property connectSet() As pscMitaConnect.CMitaConnect
		Get
			Return mitaConnect
		End Get
		Set(ByVal Value As pscMitaConnect.CMitaConnect)
			mitaConnect = Value
			Dim errorMessage As String = dbTest(Value.dbConnectString)
			If errorMessage = "" Then
				dbOpenOdbc()
				dbCloseOdbc()
				mitaData.connectionOK = True
			End If
		End Set
	End Property
	Public Property errorRecords() As Byte()
		Get
			Dim i As Integer
			Dim j As Integer
			Dim byt() As Byte
			Dim cnt As Integer = 0
			ReDim byt(mvarSAPErrRec.Length * mvarSAPErrRec(0).Length - 1)
			For i = 0 To UBound(mvarSAPErrRec)
				For j = 1 To mvarSAPErrRec(i).Length
					byt(cnt) = CByte(Asc(Mid$(mvarSAPErrRec(i), j, 1)))
					cnt = cnt + 1
				Next
			Next
			Return byt
		End Get
		Set(ByVal Value As Byte())
			mvarSAPErrRec = Nothing
		End Set
	End Property
End Class