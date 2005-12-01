Public Class CMitaTables
	Private Const mvarTableCustCombis As String = "psc§custcombis"
	Private Const mvarTableCustFields As String = "psc§custfields"
	Private Const mvarTableCustQuery As String = "psc§custquery"
	Private Const mvarTableCustTables As String = "psc§custmvarTables"
	Private Const mvarTableEventControl As String = "psc§eventcontrol"
	Private Const mvarTableEventLog As String = "psc§eventlog"
	Private Const mvarTableOnline As String = "psc§online"
	Private Const mvarTableOrderControl As String = "psc§ordercontrol"
	Private Const mvarTableOrderData As String = "psc§orderdata"
	Private Const mvarTableErrorControl As String = "psc§errorcontrol"
	Private Const mvarTableErrorData As String = "psc§errordata"
	Private Const mvarTableReportControl As String = "psc§reportcontrol"
	Private Const mvarTableSapSystems As String = "pscsapsystems"
	Private Const mvarTableStructFields As String = "psc§structfields"
	Private Const mvarTableStructures As String = "psc§structures"
	Private Const mvarTableSapVersions As String = "pscsapversions"
	Private Const mvarTableSapPool As String = "psc§sappool"
	Private mvarSapSystemID As Char = "1"
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
	Public ReadOnly Property tableSapVersions() As String
		Get
			Return mvarTableSapVersions
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
	Public ReadOnly Property tableSapSystems() As String
		Get
			Return mvarTableSapSystems
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
	Public ReadOnly Property tableCustCombis() As String
		Get
			Return mvarTableCustCombis.Replace("§"c, mvarSapSystemID)
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
	Public WriteOnly Property sapSystemid() As Integer
		Set(ByVal Value As Integer)
			mvarSapSystemID = CChar(CStr(Value))
		End Set
	End Property
End Class
