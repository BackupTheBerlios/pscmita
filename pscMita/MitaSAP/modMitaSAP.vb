Option Strict On
Option Explicit On 
Imports VB = Microsoft.VisualBasic
Imports pscMitaDef.CMitaDef
Imports MitaSAPDefClass.CMitaSAPDef
Imports System.Text
Module modMitaSAP
	Public sapVersion As Integer
	' globals for SAP data exchange
	Public sapParameterImport() As RFC_PARAMETER
	Public numberParameterImport As Integer
	Public sapParameterExport() As RFC_PARAMETER
	Public numberParameterExport As Integer
	Public sapTables() As RFC_TABLE
	Public numberTables As Integer
	Public SAPRecLen() As Integer

    Public MyError As RFC_ERROR_INFO

	' param variables
	Public xException As String

	' Global temp return value from RFC calls
	Public rc As Integer
	' public Register server handle
	Public hRFCServer As Integer = 0
	Public hRFCClient As Integer = 0
	Public hParameterSpace As Integer


	Public iClearScreen As String
	Public iFlgMotivChangeable As String
	Public iFlgTechSystemOrg As String

	Public Function tablesClose() As Integer
		numberTables = numberTables + 1
		ReDim Preserve sapTables(numberTables)
		sapTables(numberTables).name = ""
		Return numberTables
	End Function
	Public Function tablesFill(ByVal Index As Integer, ByRef DstTbl As sapStructure) As Integer
		Dim crow As Integer
		For crow = 1 To VB.UBound(DstTbl.csData) + 1
			If ItAppLine(sapTables(Index).ithandle) = 0 Then GoTo haserr
			If ItPutLine(sapTables(Index).ithandle, crow, DstTbl.csData(crow - 1)) <> 0 Then GoTo haserr
		Next crow
		Return crow - 1
hasErr:
		fillError()
		Return -2
	End Function

	Public Function tablesRead(ByVal Index As Integer, ByRef DstTbl() As String) As Integer
		Dim crow As Integer
		Dim tblString As String
		Dim cnt As Integer
		Dim tableHandle As Integer = RfcGetTableHandle(hParameterSpace, Index)
		If tableHandle = 0 Then Return -1
		ReDim DstTbl(0)
		cnt = ItFill(tableHandle)
		For crow = 1 To cnt
			If ItGupLine(tableHandle, crow) = 0 Then GoTo hasErr
			tblString = New String(VB.Chr(0), sapTables(Index).leng)
			If ItCpyLine(tableHandle, crow, tblString) <> 0 Then GoTo haserr
			ReDim Preserve DstTbl(crow - 1)
			DstTbl(crow - 1) = tblString
		Next crow
		Return crow - 1
hasErr:
		fillError()
		Return -2
	End Function

	Public Sub tablesFree()
		Dim i As Integer
		For i = 0 To UBound(sapTables)
			If sapTables(i).ithandle > 0 Then
				ItDelete(sapTables(i).ithandle)
				sapTables(i).ithandle = 0
			End If
		Next i
	End Sub

	Public Sub CloseRFCs()
		If hRFCServer <> 0 Then RfcClose(hRFCServer)
		hRFCServer = 0
		If hRFCClient <> 0 Then RfcClose(hRFCClient)
		hRFCClient = 0
	End Sub

	Public Sub tablesInit()
		ReDim sapTables(0)
		numberTables = -1
	End Sub

	Public Sub importInitParameter()
		ReDim sapParameterImport(0)
		numberParameterImport = -1
	End Sub
	Public Function importAddParameter(ByRef SapNam As String, ByVal Tp As Integer, ByVal Ln As Integer, ByRef addr As String) As Integer
		numberParameterImport = numberParameterImport + 1
		ReDim Preserve sapParameterImport(numberParameterImport)
		sapParameterImport(numberParameterImport).name = SapNam
		sapParameterImport(numberParameterImport).nlen = VB.Len(SapNam)
		sapParameterImport(numberParameterImport).type = Tp
		sapParameterImport(numberParameterImport).leng = Ln
		sapParameterImport(numberParameterImport).addr = addr
		Return RfcAddImportParam(hParameterSpace, _
		numberParameterImport, _
		sapParameterImport(numberParameterImport).name, _
		sapParameterImport(numberParameterImport).nlen, _
		sapParameterImport(numberParameterImport).type, _
		sapParameterImport(numberParameterImport).leng, _
		sapParameterImport(numberParameterImport).addr)
	End Function
	Public Function importDefineParameter(ByRef SapNam As String, ByVal Tp As Integer, ByVal Ln As Integer) As Integer
		numberParameterImport = numberParameterImport + 1
		ReDim Preserve sapParameterImport(numberParameterImport)
		sapParameterImport(numberParameterImport).name = SapNam
		sapParameterImport(numberParameterImport).nlen = VB.Len(SapNam)
		sapParameterImport(numberParameterImport).type = Tp
		sapParameterImport(numberParameterImport).leng = Ln
		Return RfcDefineImportParam(hParameterSpace, _
		numberParameterImport, _
		sapParameterImport(numberParameterImport).name, _
		sapParameterImport(numberParameterImport).nlen, _
		sapParameterImport(numberParameterImport).type, _
		sapParameterImport(numberParameterImport).leng)
	End Function
	Public Function importCloseParameter() As Integer
		importAddParameter("", 0, 0, "")
		Return numberParameterImport
	End Function
    Public Function importReadString(ByRef Index As Integer, ByVal ln As Integer) As String
        Dim a As String = New String(VB.Chr(0), sapParameterImport(Index).leng)
        Dim rc As Integer
        rc = RfcGetImportParam(hParameterSpace, Index, a)
        Return a
    End Function

    Public Sub exportInitParameter()
        ReDim sapParameterExport(0)
        numberParameterExport = -1
    End Sub
    Public Function exportAddParameter(ByRef SapNam As String, ByVal Tp As Integer, ByVal Ln As Integer, ByRef value As String) As Integer
        numberParameterExport = numberParameterExport + 1
        ReDim Preserve sapParameterExport(numberParameterExport)
        sapParameterExport(numberParameterExport).name = SapNam
        sapParameterExport(numberParameterExport).nlen = VB.Len(SapNam)
        sapParameterExport(numberParameterExport).type = Tp
        sapParameterExport(numberParameterExport).leng = Ln
        Return RfcAddExportString(hParameterSpace, _
        numberParameterExport, _
        sapParameterExport(numberParameterExport).name, _
        sapParameterExport(numberParameterExport).nlen, _
        sapParameterExport(numberParameterExport).type, _
        sapParameterExport(numberParameterExport).leng, _
        value)
    End Function
    Public Function exportCloseParameter() As Integer
        exportAddParameter("", 0, 0, "")
        Return numberParameterExport
    End Function
    Public Function exportDataBuild(ByRef action As sapFrame) As Integer
        Dim RfcRc As Integer
        Dim i As Integer
        RfcRc = allocParamSpace(action)
        If RfcRc <> RFC_RC.RFC_OK Then Return RfcRc
        exportInitParameter()
        For i = 0 To action.cfCountExport
            action.cfExportParameter(i).csLength = 313
            exportAddParameter(action.cfExportParameter(i).csStructure, _
            vTYPC, _
            action.cfExportParameter(i).csLength, _
            action.cfExportParameter(i).csData(0))
        Next
        exportCloseParameter()
    End Function
    Public Function importDataBuild(ByRef action As sapFrame) As Integer
        Dim RfcRc As Integer
        Dim i As Integer
        Dim tmp As Integer
        RfcRc = allocParamSpace(action)
        If RfcRc <> RFC_RC.RFC_OK Then Return RfcRc
        importInitParameter()
        For i = 0 To action.cfCountImport
            action.cfImportParameter(i).csLength = 313
            action.cfImportParameter(i).csData(0) = Space$(action.cfImportParameter(i).csLength)
            rc = importDefineParameter(action.cfImportParameter(i).csStructure, _
            vTYPC, _
            action.cfImportParameter(i).csLength)
        Next
        importCloseParameter()
        importDataBuild = RFC_RC.RFC_OK
exx:
        Exit Function

hasErr:
        importDataBuild = RFC_RC.RFC_FAILURE
        GoTo exx
    End Function
    Public Function allocParamSpace(ByRef action As sapFrame) As Integer
        ' allocate param space
        hParameterSpace = RfcAllocParamSpace(action.cfCountExport + 1, action.cfCountImport + 1, action.cfCountTables + 1)
        If hParameterSpace = 0 Then
            fillError()
            If hRFCClient <> 0 Then
                RfcClose(hRFCClient)
                hRFCClient = 0
            End If
            Return RFC_RC.RFC_FAILURE    'RFC failure
        End If
        Return RFC_RC.RFC_OK
    End Function
    Public Function createSapFrame(ByVal rfcClass As Integer, ByRef cStruct() As STRUCTSTRUCT, ByVal cCount As Integer) As sapFrame
        Dim fram As New sapFrame
        fram.cfCountExport = -1
        fram.cfCountImport = -1
        fram.cfCountTables = -1
        fillFrame(fram, cStruct, cCount)
        Return fram
    End Function
    Private Function fillFrame(ByRef fram As sapFrame, ByRef cStruct() As STRUCTSTRUCT, ByVal cCount As Integer) As Integer
        Dim RfcRc As Integer
        Dim strucName As String
        Dim i As Integer
        Dim tmp As Integer
        For i = 0 To cCount
            If Not IsNothing(cStruct(i).sName) Then
                Select Case cStruct(i).sType
                    Case "I"c
                        fram.cfCountImport = fram.cfCountImport + 1
                        tmp = fram.cfCountImport
                        ReDim Preserve fram.cfImportParameter(tmp)
                        fram.cfImportParameter(tmp).csCount = -1
                        fram.cfImportParameter(tmp).csStructure = cStruct(i).sData
                        fram.cfImportParameter(tmp).csCount = fram.cfImportParameter(tmp).csCount + 1
                        ReDim Preserve fram.cfImportParameter(tmp).csData(fram.cfImportParameter(tmp).csCount)
                        fram.cfImportParameter(tmp).csLength = cStruct(i).sLength
                        fram.cfImportParameter(tmp).csType = clientSapRfcType.rfcImportParameter
                    Case "E"c
                        fram.cfCountExport = fram.cfCountExport + 1
                        tmp = fram.cfCountExport
                        ReDim Preserve fram.cfExportParameter(tmp)
                        fram.cfExportParameter(tmp).csCount = -1
                        fram.cfExportParameter(tmp).csStructure = cStruct(i).sData
                        fram.cfExportParameter(tmp).csCount = fram.cfExportParameter(tmp).csCount + 1
                        ReDim Preserve fram.cfExportParameter(tmp).csData(fram.cfExportParameter(tmp).csCount)
                        fram.cfExportParameter(tmp).csLength = cStruct(i).sLength
                        fram.cfExportParameter(tmp).csType = clientSapRfcType.rfcExportParameter
                    Case "B"c
                        fram.cfCountImport = fram.cfCountImport + 1
                        tmp = fram.cfCountImport
                        ReDim Preserve fram.cfImportParameter(tmp)
                        fram.cfImportParameter(tmp).csCount = -1
                        fram.cfImportParameter(tmp).csStructure = cStruct(i).sData
                        fram.cfImportParameter(tmp).csCount = fram.cfImportParameter(tmp).csCount + 1
                        ReDim Preserve fram.cfImportParameter(tmp).csData(fram.cfImportParameter(tmp).csCount)
                        fram.cfImportParameter(tmp).csLength = cStruct(i).sLength
                        fram.cfImportParameter(tmp).csType = clientSapRfcType.rfcImportParameter
                        fram.cfCountExport = fram.cfCountExport + 1
                        tmp = fram.cfCountExport
                        ReDim Preserve fram.cfExportParameter(tmp)
                        fram.cfExportParameter(tmp).csCount = -1
                        fram.cfExportParameter(tmp).csStructure = cStruct(i).sData
                        fram.cfExportParameter(tmp).csCount = fram.cfExportParameter(tmp).csCount + 1
                        ReDim Preserve fram.cfExportParameter(tmp).csData(fram.cfExportParameter(tmp).csCount)
                        fram.cfExportParameter(tmp).csLength = cStruct(i).sLength
                        fram.cfExportParameter(tmp).csType = clientSapRfcType.rfcExportParameter
                    Case "T"c
                        fram.cfCountTables = fram.cfCountTables + 1
                        ReDim Preserve fram.cfDataTables(fram.cfCountTables)
                        fram.cfDataTables(fram.cfCountTables).csCount = -1
                        fram.cfDataTables(fram.cfCountTables).csStructure = cStruct(i).sData
                        fram.cfDataTables(fram.cfCountTables).csCount = fram.cfDataTables(fram.cfCountTables).csCount + 1
                        ReDim Preserve fram.cfDataTables(fram.cfCountTables).csData(fram.cfDataTables(fram.cfCountTables).csCount)
                        fram.cfDataTables(fram.cfCountTables).csLength = cStruct(i).sLength
                        fram.cfDataTables(fram.cfCountTables).csType = clientSapRfcType.rfcTables
                End Select
            End If
        Next
    End Function
    Public Sub fillError()
        Try
            RfcLastError(MyError)
        Catch
            MyError.message = Err.Description
        End Try
    End Sub
End Module