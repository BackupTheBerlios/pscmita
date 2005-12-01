Imports pscMitaDef.CMitaDef
Public Class CMitaData
	Public structureCount As Integer
	Public structures() As STRUCTSTRUCT
	Public oldStruct As String
	Public isMiniSAP As Boolean = False
	Public doTrace As Integer = 0
	Public rfcName As String
	Public mitaApplication As String
	Public registryApplication As String
	Public errorDescription As String
	Public applicationForm As System.Windows.Forms.Form
	Public popupTime As Short = 0
	Public caption As String

	Public tablesCount As Integer
	Public tables() As TABLESTRUCT

	Public combiCount As Integer
	Public combis() As COMBISTRUCT

	Public createHost As String
	Public createID As Integer = -1
	Public newID As Integer
	Public tmpID As Integer
	Public createUser As String

	Public custFieldCount As Integer
	Public custFields() As FIELDSTRUCT
	Public structFieldCount As Integer
	Public structFields() As FIELDSTRUCT

	Public eventCount As Integer
	Public mEvents() As EVENTSTRUCT

	Public connectionOK As Boolean
	Public processID As Integer
	Public commandLine As String

	Public createOwner As String = ""
	Public createDatabase As String = ""
	Public createType As String = "Oracle"

End Class
