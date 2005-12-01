Imports pscMitaDef.CMitaDef
Imports MitaSAPDefClass.CMitaSAPDef
'<ComClass(SapClient.ClassId, SapClient.InterfaceId, SapClient.EventsId)> _
Public Class CSapClient

	'#Region "COM GUIDs"
	'	' These  GUIDs provide the COM identity for this class 
	'	' and its COM interfaces. If you change them, existing 
	'	' clients will no longer be able to access the class.
	'	Public Const ClassId As String = "1E3886C0-CBF7-4299-905D-59F7E88CAC0A"
	'	Public Const InterfaceId As String = "BD99BC28-ECDE-4149-9F21-4D646072CF62"
	'	Public Const EventsId As String = "701A4C23-E76E-4F6A-BD6B-F5878E101E5C"
	'#End Region

	Public Sub New()
		MyBase.New()
        MyError = New RFC_ERROR_INFO
    End Sub
    Public Function clientSapSend(ByVal action As sapFrame) As Boolean
        Dim result As Boolean
        Dim mode As RFC_MODE
        Dim dest As String
        Dim xerr As RFC_ERROR_INFO_EX
        mode = RFC_MODE.RFC_MODE_R3ONLY
        dest = ""
        hRFCClient = RfcOpenExt(dest, _
        mode, _
        action.cfServer, _
        CInt(action.cfSystem), _
        action.cfGateway, _
        action.cfService, _
        action.cfMandant, _
        action.cfUser, _
        action.cfPassword, _
        action.cfLanguage, _
        action.cfTrace)
        If hRFCClient = 0 Then
            fillError()
            Return False
        End If
        result = (rfcClientSend(action) = RFC_RC.RFC_OK)
        If hRFCClient <> 0 Then
            RfcClose(hRFCClient)
            hRFCClient = 0
        End If
        Return result
    End Function
    Private Function rfcClientSend(ByRef action As sapFrame) As Integer
        Dim RfcRc As Integer
        Dim i As Integer
        Dim funcName As String
        funcName = action.cfFunction
        If importDataBuild(action) <> RFC_RC.RFC_OK Then GoTo hasErr
        If tablesBuildClient(action) <> RFC_RC.RFC_OK Then GoTo hasErr
        For i = 0 To action.cfCountTables
            rc = tablesFill(i, action.cfDataTables(i))
        Next i
        If rc = -2 Then GoTo hasErr
        xException = New String(Chr(0), 256)
        RfcRc = RfcCallReceiveExt(hRFCClient, hParameterSpace, action.cfFunction, xException)

        If RfcRc <> RFC_RC.RFC_OK Then GoTo hasErr
        rfcClientSend = RfcRc
exx:
        rc = RfcFreeParamSpace(hParameterSpace)
        Exit Function

hasErr:
        rfcClientSend = RFC_RC.RFC_FAILURE
        fillError()
        GoTo exx
    End Function
    Public Function getLastError() As String
        fillError()
        Return MyError.message
    End Function
    Public Function createFrame(ByVal rfcClass As Integer, ByRef cStruct() As STRUCTSTRUCT, ByVal cCount As Integer) As sapFrame
        Return createSapFrame(rfcClass, cStruct, cCount)
    End Function
    Private Function tablesBuildClient(ByRef action As sapFrame) As Integer
        Dim i As Integer
        tablesInit()
        For i = 0 To action.cfCountTables
            action.cfDataTables(i).csLength = action.cfDataTables(i).csLength
            rc = tablesAddClient(action.cfDataTables(i).csStructure, vTYPC, action.cfDataTables(i).csLength)
        Next
        tablesClose()
    End Function
    Private Function tablesAddClient(ByRef SapNam As String, ByVal Tp As Integer, ByVal Ln As Integer) As Integer
        numberTables = numberTables + 1
        ReDim Preserve sapTables(numberTables)
        sapTables(numberTables).name = SapNam
        sapTables(numberTables).nlen = Len(SapNam)
        sapTables(numberTables).type = Tp
        sapTables(numberTables).leng = Ln
        sapTables(numberTables).ithandle = ItCreate(sapTables(numberTables).name, Ln, 0, 0)
        Return RfcAddTable(hParameterSpace, numberTables, sapTables(numberTables).name, sapTables(numberTables).nlen, sapTables(numberTables).type, sapTables(numberTables).leng, sapTables(numberTables).ithandle)
    End Function
End Class


