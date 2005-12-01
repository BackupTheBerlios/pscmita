Imports System.Data.Odbc
Imports pscMitaDef.CMitaDef
Public Class CMitaSapSystem
	Public Const tableCustCombis As String = "psccustcombis"
	Public Const tableSapSystems As String = "pscsapsystems"
	Public Const tableSapVersions As String = "pscsapversions"

	Private Const mvarTableCustFields As String = "psc§custfields"
	Private Const mvarTableCustQuery As String = "psc§custquery"
	Private Const mvarTableCustTables As String = "psc§custtables"
	Private Const mvarTableEventControl As String = "psc§eventcontrol"
	Private Const mvarTableEventLog As String = "psc§eventlog"
	Private Const mvarTableOnline As String = "psc§online"
	Private Const mvarTableOrderControl As String = "psc§ordercontrol"
	Private Const mvarTableOrderData As String = "psc§orderdata"
	Private Const mvarTableErrorControl As String = "psc§errorcontrol"
	Private Const mvarTableErrorData As String = "psc§errordata"
	Private Const mvarTableReportControl As String = "psc§reportcontrol"
	Private Const mvarTableStructFields As String = "psc§structfields"
	Private Const mvarTableStructures As String = "psc§structures"
	Private Const mvarTableSapPool As String = "psc§sappool"

	'Public mitaTableSpace As String
	Public rfcType As rfcClass
	Private mvarSapSystemIDString As String
	Public sapSystemNAME As String
	Public sapSystemSAPGATEWAY As String
	Public sapSystemSAPSERVICE As String
	Public sapSystemSAPID As String
	Public sapSystemSAPIDCLT As String
	Public sapSystemSAPUSER As String
	Public sapSystemSAPOWNER As String
	Public sapSystemSAPSYSTEM As String
	Public sapSystemSAPSERVER As String
	Public sapSystemSAPCLIENT As String
	Public sapSystemDATABASE As String
	Public sapSystemDATABASETYPE As String
	Public sapSystemDATABASEUSER As String
	Public sapSystemDATABASEPWD As String
	Public sapSystemVERSIONID As Integer
	Public sapSystemVERSIONNAME As String
	Public sapSystemSUBVERSION As String

	Public runType As String

	Private mvarErrorDescription As String
	Private mvarSapPool As String = mvarTableSapPool
	Private mvarSapSystemID As Char = "1"
	Public Property sapSystemId() As String
		Get
			Return mvarSapSystemIDString
		End Get
		Set(ByVal Value As String)
			mvarSapSystemIDString = Value
			mvarSapSystemID = CChar(CStr(Value))
		End Set
	End Property
	Public Property sapPool() As String
		Get
			Return mvarSapPool.Replace("§"c, mvarSapSystemID)
		End Get
		Set(ByVal Value As String)
			mvarSapPool = Value
		End Set
	End Property

	Public ReadOnly Property tableStructFields() As String
		Get
			Return mvarTableStructFields.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableStructures() As String
		Get
			Return mvarTableStructures.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableSapPool() As String
		Get
			Return mvarTableSapPool.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableErrorControl() As String
		Get
			Return mvarTableErrorControl.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableErrorData() As String
		Get
			Return mvarTableErrorData.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableReportControl() As String
		Get
			Return mvarTableReportControl.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableOrderControl() As String
		Get
			Return mvarTableOrderControl.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableOrderData() As String
		Get
			Return mvarTableOrderData.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableCustFields() As String
		Get
			Return mvarTableCustFields.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableCustQuery() As String
		Get
			Return mvarTableCustQuery.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableCustTables() As String
		Get
			Return mvarTableCustTables.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableEventControl() As String
		Get
			Return mvarTableEventControl.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableEventLog() As String
		Get
			Return mvarTableEventLog.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property tableOnline() As String
		Get
			Return mvarTableOnline.Replace("§"c, mvarSapSystemID)
		End Get
	End Property
	Public ReadOnly Property connectString() As String
		Get
			Return "DSN=" & sapSystemDATABASE & ";UID=" & sapSystemDATABASEUSER & ";PWD=" & sapSystemDATABASEPWD & ";"
		End Get
	End Property
End Class
