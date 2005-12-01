Imports pscMitaDef.CMitaDef
Imports MitaSAPDefClass.CMitaSAPDef
'<ComClass(SapServer.ClassId, SapServer.InterfaceId, SapServer.EventsId)> _
Public Class CSapServer

	Dim mvarOrderFrame As sapFrame

	'#Region "COM GUIDs"
	'	' These  GUIDs provide the COM identity for this class 
	'	' and its COM interfaces. If you change them, existing 
	'	' clients will no longer be able to access the class.
	'	Public Const ClassId As String = "2EC507D6-ED46-4D25-A17A-BD6ED6CC817C"
	'	Public Const InterfaceId As String = "21B4253A-E130-4A4A-BAD9-785A80D6372A"
	'	Public Const EventsId As String = "0F2A5A63-48D5-478D-B77B-ECCB83FE3DF4"
	'#End Region

	' A creatable COM class must have a Public Sub New() 
	' with no parameters, otherwise, the class will not be 
	' registered in the COM registry and cannot be created 
	' via CreateObject.
	Public Sub New()
        MyBase.New()
        MyError = New RFC_ERROR_INFO
	End Sub

	Public Function getLastError() As String
		fillError()
		Return MyError.message
	End Function
	Public Function orderGetNext() As Boolean
		Dim rfcrc As RFC_RC
		If IsNothing(hRFCServer) Then Return False
		rfcrc = RfcListen(hRFCServer)
		Select Case rfcrc
			Case RFC_RC.RFC_OK
				rfcrc = RfcDispatch(hRFCServer)
				If rfcrc = RFC_RC.RFC_OK Then Return True
				fillError()
				Return False
			Case RFC_RC.RFC_RETRY
				Return False
			Case RFC_RC.RFC_FAILURE
				fillError()
				Return False
		End Select
	End Function
	Private Function sapCloseInput() As Boolean
		If hRFCServer <> 0 Then
			RfcClose(hRFCServer)
			hRFCServer = 0
		End If
		Return True
	End Function
	Public Function sapInitInput(ByVal action As sapFrame) As Integer
		Dim Cmd$ = "-a" & action.cfHost & "." & action.cfProgram & " -g" & action.cfServer & " -x" & action.cfService
		If action.cfTrace <> 0 Then
			Cmd$ = Cmd$ & " -t"
		End If
		REM register myself as sap server
		hRFCServer = RfcAcceptExt(Cmd$)
		If hRFCServer = 0 Then
			fillError()
			GoTo exx
		End If
		REM tell sap about my function addresses
		Select Case action.cfTyp
			Case rfcClass.rfcOrder
				rc = RfcInstallFunctionExt(hRFCServer, action.cfFunction, AddressOf srv_isp_adprodorder_save, "srv_isp_adprodorder_save")
			Case Else
				GoTo exx
		End Select
		If rc <> RFC_RC.RFC_OK Then GoTo exx
		mvarOrderFrame = action
		sapInitInput = hRFCServer
		Exit Function
exx:
		sapInitInput = 0
	End Function
	Private Function srv_isp_adprodorder_save(ByVal hRFCServer As Integer) As Integer
		Dim RfcRc As Integer
		Dim i As Integer
		Dim funcName As String
		funcName = mvarOrderFrame.cfFunction
		If importDataBuild(mvarOrderFrame) <> RFC_RC.RFC_OK Then GoTo hasErr
		If tablesBuildServer(mvarOrderFrame) <> RFC_RC.RFC_OK Then GoTo hasErr
		RfcRc = RfcGetDataExt(hRFCServer, hParameterSpace)
		If RfcRc <> 0 Then GoTo hasErr
		'		REM read imported parameters
		mvarOrderFrame.cfImportParameter(0).csData(0) = importReadString(0, mvarOrderFrame.cfImportParameter(0).csLength)
		'		REM Read imported tables
		For i = 0 To mvarOrderFrame.cfCountTables
			rc = tablesRead(i, mvarOrderFrame.cfDataTables(i).csData)
			If rc = -2 Then GoTo hasErr
		Next
        If exportDataBuild(mvarOrderFrame) <> RFC_RC.RFC_OK Then GoTo hasErr
        mvarOrderFrame.cfExportParameter(0).csData(0) = "TEST"
        If RfcRc = RFC_RC.RFC_OK Then RfcRc = RfcSendDataExt(hRFCServer, hParameterSpace) 'RFC_OK

        If RfcRc <> RFC_RC.RFC_OK Then GoTo hasErr
        srv_isp_adprodorder_save = RfcRc
exx:
        rc = RfcFreeParamSpace(hParameterSpace)
        tablesFree()
        Exit Function

hasErr:
        srv_isp_adprodorder_save = RFC_RC.RFC_FAILURE
        fillError()
        GoTo exx
	End Function
	Public Function createFrame(ByVal add As String, ByVal rfcClass As Integer, ByRef cStruct() As STRUCTSTRUCT, ByVal cCount As Integer) As sapFrame
		Return createSapFrame(rfcClass, cStruct, cCount)
	End Function

	Protected Overrides Sub Finalize()
		sapCloseInput()
		MyBase.Finalize()
	End Sub
	Private Function tablesAddServer(ByRef SapNam As String, ByVal Tp As Integer, ByVal Ln As Integer) As Integer
		numberTables = numberTables + 1
		ReDim Preserve sapTables(numberTables)
		sapTables(numberTables).name = SapNam
		sapTables(numberTables).nlen = Len(SapNam)
		sapTables(numberTables).type = Tp
		Return RfcAddTable(hParameterSpace, numberTables, sapTables(numberTables).name, sapTables(numberTables).nlen, sapTables(numberTables).type, sapTables(numberTables).leng, sapTables(numberTables).ithandle)
	End Function
	Private Function tablesBuildServer(ByRef action As sapFrame) As Integer
		Dim i As Integer
		tablesInit()
		For i = 0 To action.cfCountTables
			action.cfDataTables(i).csLength = action.cfDataTables(i).csLength
			rc = tablesAddServer(action.cfDataTables(i).csStructure, vTYPC, action.cfDataTables(i).csLength)
		Next
		tablesClose()
	End Function
End Class


