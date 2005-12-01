Imports pscMitaDef.CMitaDef
Module pscMitaSharedData
	Public sapPool As String = "pscsappool"

	Public sapSystemID As Integer
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
	Public sapSystemDATABASEUSER As String
	Public sapSystemDATABASEPWD As String
	Public sapSystemVERSIONNAME As String
	Public sapSystemVERSIONID As Integer
	Public sapSystemSUBVERSION As String

	Public structureCount As Integer
	Public structures() As STRUCTSTRUCT
	Public oldStruct As String
	Public isMiniSAP As Boolean = False
	Public doTrace As Integer = 0
	Public rfcType As Integer
	Public rfcName As String
	Public mitaApplication As String
	Public errorDescription As String
	Public frmMitaMsgInst As pscMitaMsg.CMitaMsg
	Public popupTime As Short = 0


	Public tablesCount As Integer
	Public tables() As TABLESTRUCT

	Public combiCount As Integer
	Public combis() As COMBISTRUCT

	Public createHost As String
	Public createID As Integer = 0
	Public newID As Integer
	Public tmpID As Integer
	Public createUser As String

	Public custFieldCount As Integer
	Public custFields() As FIELDSTRUCT
	Public structFieldCount As Integer
	Public structFields() As FIELDSTRUCT

	Public sapName As String
	Public runType As String


	Public eventCount As Integer
	Public mEvents() As EVENTSTRUCT

End Module
