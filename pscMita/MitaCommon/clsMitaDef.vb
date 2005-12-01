Option Strict On
Option Explicit On 

Public Class CMitaDef

	' IMPORTANT!
	' this module should make the programmer's work easier
	' Do not change ANYTHING in this module!
	' if you do so, the PscMita modules may not work anymore!

	Public Const StrRecord As Short = 6

	Public Structure COMBISTRUCT
		Dim cName As String
		Dim cEntry As String
	End Structure

	Public Structure EVENTACTION
		Dim aName As String
		Dim aRecno As Integer
		Dim aTodo As String
		Dim aComment As String
	End Structure

	Public Structure EVENTSTRUCT
		Dim sName As String
		Dim sCount As Integer
		Dim sActions() As EVENTACTION
		Dim sComment As String
	End Structure
	Public Structure TABLESTRUCT
		Dim tName As String
		Dim tCount As Integer
		Dim tEntries() As PAIRSTRUCT
	End Structure

	Public Enum rfcClass
		rfcOrder = 1
		rfcDesign
		rfcPosition
		rfcError
		rfcStatRead
		rfcStatWrite
		rfcTextRead
		rfcTextWrite
		rfcMax		' must be last entry
	End Enum
	Public Enum mitaLogClass
		logNormal
		logError
		logSQL
	End Enum
	Public Enum mitaSqlClass
		classError = 1
		classInternal = 2
		classOrder = 4
		classAd = 8
		classAll = mitaSqlClass.classError + mitaSqlClass.classInternal + mitaSqlClass.classOrder + mitaSqlClass.classAd
	End Enum

	Public Const pakINDEX As Integer = 0
	Public Const papIndex As Integer = 1
	Public Const bpzINDEX As Integer = 2
	Public Const iszINDEX As Integer = 3
	Public Const moINDEX As Integer = 4
	Public Const blzINDEX As Integer = 5
	Public Const psINDEX As Integer = 6
	Public Const plzINDEX As Integer = 7
	Public Const plzaINDEX As Integer = 8
	Public Const statINDEX As Integer = 9
	Public Const txtINDEX As Integer = 10
	Public Const psiINDEX As Integer = 11
	Public Const errINDEX As Integer = 12

	Public Enum mitaEventCodes
		' internal errors
		errorProgrammer = 1
		errorUserIni
		errorSAPData
		errorSAPVersion
		errorSAPConnection
		errorSAPReference
		errorFileSystem
		errorDataBase
		errorNoOpenOrder
		errorDatabaseSequence
		errorDatabaseConnect
		errorNoActualSelect
		errorNoRowsSelected
		errorInvalidInput
		errorNoHost
		errorLoop
		errorRequery
		errorRequeryYes
		errorRequeryNo
		errorTrialsExceeded
		errorTrialsPossible
		' errors raised from application
		userSqlException
		' events raised from application
		programOrderRead		' MUST be first non-error event
		programOrderWritten
		programStart
		programEnd
		userOrderSuccess
		userOrderFailure
		userMotivSuccess
		userLogSQL
		userException1
		userException2
		userMessage
		codesMax		' MUST be last entry!
	End Enum

	Public Enum mitaErrorCodes
		orderOK = 0
		orderSapDataError = 1
		orderSapOldVersion = 2
		orderUserSQLException = 4
		orderNoPublishDate = 8
		orderWriteDBProblem = 16
	End Enum

	Public Enum mitaEventReturnCodes
		returnOk = 0
		requeryRequest
		requeryExceeded
		debugBreak
	End Enum

	Public Structure SQLLOG
		Dim sName As String
		Dim sTemplate As String
		Dim sResult As String
		Dim sNumberAd As String
		Dim sVersionAd As String
		Dim sPaper As String
		Dim sClass As mitaSqlClass
		Dim sError As String
	End Structure

	Public Structure STRINGBOOL
		Dim sbContent As String
		Dim sbUsed As Boolean
	End Structure

	Public Structure MITATABLE
		Dim mtExt As String
		Dim mtName As String
		Dim mtLength As Integer
		Dim mtRecords As STRINGBOOL()
		Dim mtCount As Integer
	End Structure

	Public Structure MITAFIELD
		Dim mfRecord As Integer
		Dim mfFirst As Integer
		Dim mfLength As Integer
		Dim mfTyp As String
		Dim mfLabel As String
	End Structure

	Public Structure STRUCTSTRUCT
		Dim sFunction As String
		Dim sIndex As Integer
		Dim sName As String
		Dim sData As String
		Dim sLength As Integer
		Dim sLevel As Integer
		Dim sType As Char
	End Structure

	Public Structure PAIRSTRUCT
		Dim pLeft As String
		Dim pRight As String
	End Structure

	Public Structure FIELDSTRUCT
		Dim fName As String
		Dim fFirst As Integer
		Dim fLength As Integer
		Dim fSaved As Boolean
		Dim fStructure As String
		Dim fStructureField As String
		Dim fType As String
		Dim fIndex As Integer
		Dim fRemoved As Boolean
		Dim fLevel As Integer
	End Structure
	Public Structure custTableArray
		Dim name As String
		Dim tableLeftLeft() As String
		Dim tableLeftRight() As String
		Dim tableRightLeft() As String
		Dim tableRightRight() As String
	End Structure
	Public Structure custTable
		Dim tleftleft As String
		Dim trightleft As String
		Dim tleftright As String
		Dim trightright As String
	End Structure
	Enum clientSapRfcType
		rfcTables
		rfcImportParameter
		rfcExportParameter
	End Enum
	Public Structure sapStructure
		Dim csType As clientSapRfcType
		Dim csStructure As String
		Dim csCount As Integer
		Dim csData() As String
		Dim csLength As Integer
	End Structure
	Public Structure sapFrame
		Dim cfHost As String
		Dim cfProgram As String
		Dim cfServer As String
		Dim cfSystem As String
		Dim cfGateway As String
		Dim cfService As String
		Dim cfMandant As String
		Dim cfUser As String
		Dim cfPassword As String
		Dim cfLanguage As String
		Dim cfTrace As Integer
		Dim cfFunction As String
		Dim cfTyp As rfcClass
		Dim cfCountTables As Integer
		Dim cfCountImport As Integer
		Dim cfCountExport As Integer
		Dim cfDataTables() As sapStructure
		Dim cfExportParameter() As sapStructure
		Dim cfImportParameter() As sapStructure
	End Structure

End Class