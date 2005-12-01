Option Strict On
Option Explicit On 
Imports VB = Microsoft.VisualBasic
Imports System.Runtime.InteropServices
Imports System.Text

Public Class CMitaSAPDef

	'********************************************************************************
    '------ ISP_AD_DESIGN               Aufruf Editor (Aufruf von SAP)
	'------ ISP_ADPRODORDER_SAVE        Speichern der Daten (Aufruf von SAP)
	'------ ISP_ADPRODORDER_CANCEL      Verwerfen des aktuellen Auftrags
	'------                             (Aufruf von SAP)
	'****** ISP_STATUS_UPDATE_TS        Update der Staten beider Systeme
	'------                             (von SAP initiiert) (Aufruf von SAP)
	'****** ISP_STATUS_UPDATE_ISPAM     Update der Staten beider Systeme
    '------                             (von Tech Sys initiiert)
	'****** ISP_ADPRODORDER_UPDATE      Autragsänderung vom technischen System
	'------                             (vorhandener Auftrag)
	'****** ISP_CHIFFRENR_GET           Abruf einer Chiffrenummer
	'------ ISP_POSITION                Ermitteln der Plazierung (Aufruf von SAP)
	'------ ISP_POSITION_DIALOG         Aufruf des Plazierungsprogramms (Aufruf von SAP)
	'------ ISP_TECH_SYSTEM_START       Aufruf des technischen Systems ohne
	'------                             konkreten Auftrag (Aufruf von SAP)
	'****** ISP_PRODUCTION_FINISHED     Produktionsfertigmeldung (Abstrich)
	'****** ISP_ADPROOFDATA_GET         Übergabe der Motivdatei für Ausdruck
	'****** ISP_STATUS_GET              Abfrage des Status des Partnersystems
	'****** ISP_EXT_ADPRODORDER_SAVE    Übergabe im technischen System erfasster
	'------                             Aufträge (Offline erfasst)
	'****** ISP_EXCHANGE_ERRORMESSAGES  Austausch von Fehlermeldungen
	'********************************************************************************
	'
	' --------------------------------------------------------------------------
	'
	'   Called function ISP_AD_DESIGN
	'       sapParameterExport
	'         CLEAR_SCREEN like RJHA800_CLEAR_SCR default SPACE public type  RFC_CHAR length 1 as STRING
	'         FLG_MOTIV_CHANGEABLE like TJ180_TRTYP  public type  RFC_CHAR length 1 as STRING
	'         FLG_TECH_SYSTEM_ORG like SYST_BATCH default SPACE public type  RFC_CHAR length 1 as STRING
	'       sapParameter
	'         FLG_MOTIV_CHANGED like SYST_BATCH public type  STRING length 1
	'       sapTables
	'         RJHATBLZ_ITAB structure RJHATBLZ length 46 number of fields 3
	'         RJHATISZ_ITAB structure RJHATISZ length 244 number of fields 15
	'         RJHATMO_ITAB structure RJHATMO length 883 number of fields 109
	'         RJHATSTAT_ITAB structure RJHATSTAT length 659 number of fields 71
	'         RJHATTXT_ITAB structure RJHATTXT length 163 number of fields 7
	'       exceptions
	'
	' --------------------------------------------------------------------------

	' --------------------------------------------------------------------------
	'
	'   Called function ISP_TECH_SYSTEM_START
	'       sapParameterExport
	'         X_CALL_DESIGNER like SYST_BATCH  type RFC_CHAR length 1 as STRING
	'       sapParameter
	'       sapTables
	'         RJHATBLZ_ITAB structure RJHATBLZ length 46 number of fields 3
	'         RJHATMO_ITAB structure RJHATMO length 883 number of fields 109
	'         RJHATTXT_ITAB structure RJHATTXT length 163 number of fields 7
	'       exceptions
	'
	' --------------------------------------------------------------------------

	' --------------------------------------------------------------------------
	'
	'   Called function ISP_ADPRODORDER_SAVE
	'       sapParameterExport
	'         RJHATPAK_WA structure RJHATPAK length 361 number of fields 26
	'       sapParameter
	'       sapTables
	'         RJHATBLZ_ITAB structure RJHATBLZ length 46 number of fields 3
	'         RJHATBPZ_ITAB structure RJHATBPZ length 23 number of fields 3
	'         RJHATISZ_ITAB structure RJHATISZ length 244 number of fields 15
	'         RJHATMO_ITAB structure RJHATMO length 883 number of fields 109
	'         RJHATPAP_ITAB structure RJHATPAP length 102 number of fields 16
	'         RJHATPLZA_ITAB structure RJHATPLZA length 70 number of fields 15
	'         RJHATPLZ_ITAB structure RJHATPLZ length 31 number of fields 10
	'         RJHATPS_ITAB structure RJHATPS length 156 number of fields 23
	'         RJHATSTAT_ITAB structure RJHATSTAT length 659 number of fields 71
	'         RJHATTXT_ITAB structure RJHATTXT length 163 number of fields 7
	'       exceptions
	'
	' --------------------------------------------------------------------------

	' ABAP/4 data types      ANSI C         Visual Basic    Comment
	'  Const vTYPC = 0           'RFC_CHAR       STRING $        characters
	'  Const vTYPDATE = 1        'RFC_DATE       STRING $        date (YYYYMMDD)
	'  Const vTYPP = 2           'RFC_BCD        STRING $        packed numbers
	'  Const vTYPTIME = 3        'RFC_TIME       STRING $        time (HHMMSS)
	'  Const vTYPX = 4           'RFC_BYTE       STRING $        raw data
	'  Const vTYPTABH = 5        'not used here
	'  Const vTYPNUM = 6         'RFC_NUM        STRING $        digits
	'  Const vTYPFLOAT = 7       'RFC_FLOAT      FLOAT #         floating point
	'  Const vTYPINT = 8         'RFC_INT        LONG &          4 byte integer
	'  Const vTYPINT2 = 9        'RFC_INT2       INTEGER %       2 byte integer
	'  Const vTYPINT1 = 10       'RFC_INT1       INTEGER %       1 byte integer
	'  Const vTYPB = 11          'not used here
	'  Const vTYP1 = 12          'not used here
	'  Const vTYP2 = 13          'not used here


	Public Const cRFCDISPATCH As Short = 1
	Public Const cRFCLISTEN As Short = 1

	' rfc error definitions
	Enum RFC_ERROR_GROUP
		RFC_ERROR_PROGRAM = 101		' @emem Error in RFC program
		RFC_ERROR_COMMUNICATION = 102		' @emem Error in Network  as long Communications
		RFC_ERROR_LOGON_FAILURE = 103		' @emem SAP logon error
		RFC_ERROR_SYSTEM_FAILURE = 104		' @emem e.g. SAP system exception raised
		RFC_ERROR_APPLICATION_EXCEPTION = 105		' @emem The called function module raised
		' an exception
		RFC_ERROR_RESOURCE = 106		' @emem Resource not available
		' (e.g. memory insufficient,...)
		RFC_ERROR_PROTOCOL = 107		' @emem RFC Protocol error
		RFC_ERROR_INTERNAL = 108		' @emem RFC Internal error
		RFC_ERROR_CANCELLED = 109		' @emem RFC Registered Server was
		' cancelled
		RFC_ERROR_BUSY = 110		' @emem System is busy, try later
	End Enum

	<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> Structure RFC_ERROR_INFO
		<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=33)> Public key As String
		<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> Public status As String
		<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=256)> Public message As String
		<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)> Public intstat As String
	End Structure

	<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> Structure RFC_ERROR_INFO_EX
		Dim group As RFC_ERROR_GROUP		' @field error group <t RFC_ERROR_GROUP> */
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=33)> Public key As String  ' The error group (type integer) and error key
		' RFC program can better analyze the error
		' with the error group instead of error key.
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=513)> Public Message As String
	End Structure

	Public Enum RFC_MODE
		RFC_MODE_R3ONLY
		RFC_MODE_CPIC
		RFC_MODE_VERSION3
		RFC_MODE_PARAMETER
	End Enum
	'
	'Public Type RFC_CONNOP_CPIC
	'  Gateway_HostPtr As String * 16
	'  Gateway_ServPtr As String * 16
	'End Type
	''
	'Public Type RFC_OPTIONS    ' fuer rfcOpen
	'  DestPtr As Long
	'  Mode As RFC_MODE
	'  ConnOptPtr As Long
	'  ClientPtr As Long
	'  UserPtr As Long
	'  PasswPtr As Long
	'  LanguagePtr As Long
	'  Trace As Integer
	'End Type
	''
	'Public Type RFC_OPTIONS    ' fuer rfcOpen
	'  DestPtr As String
	'  Mode As Integer
	'  ConnOptPtr As String * 32
	'  ClientPtr As String
	'  UserPtr As String
	'  PasswPtr As String
	'  LanguagePtr As String
	'  Trace As Integer
	'End Type

	' RFC return values definitions
	Enum RFC_RC
		RFC_OK		' O.K.
		RFC_FAILURE		' Error occurred
		RFC_EXCEPTION		' Exception raised
		RFC_SYS_EXCEPTION		' System exception raised, connection closed
		RFC_CALL		' Call received
		RFC_INTERNAL_COM		' Internal communication, repeat (internal use only)
		RFC_CLOSED		' Connection closed by the other side
		RFC_RETRY		' No data yet (RfcListen or RfcWaitForRequest only)
		RFC_NO_TID		' No Transaction ID available
		RFC_EXECUTED		' Function already executed
		RFC_SYNCHRONIZE		' Synchronous Call in Progress (only for Windows)
		RFC_MEMORY_INSUFFICIENT		' Memory insufficient
		RFC_VERSION_MISMATCH		' Version mismatch
		RFC_NOT_FOUND		' Function not found (internal use only)
		RFC_CALL_NOT_SUPPORTED		' This call is not supported on WINDOWS
		RFC_NOT_OWNER		' Caller does not own the specified handle
		RFC_NOT_INITIALIZED		' RFC not yet initialized
		RFC_SYSTEM_CALLED		' A system call such as RFC_PING for connection
	End Enum

	' RFC parameter type definition
	Structure RFC_PARAMETER
		Dim name As String		'name of the field  (in the interface definition of the function)
		Dim nlen As Integer		'length of the name (should be len(name))
		Dim type As Integer		'datatype of the field
		Dim leng As Integer		'length of the field in Bytes
		Dim addr As String		'address of the field to be exported or imported
	End Structure

	' RFC table type definition
	Structure RFC_TABLE
		Dim name As String		'name of the table (in the interface definition of the function)
		Dim nlen As Integer		 'length of the name (should be len(name))
		Dim type As Integer		'datatype of the lines of the table
		Dim leng As Integer		'length of a row in bytes
		Dim ithandle As Integer		'table handle (type ITAB_H)
		Dim itmode As Integer		'mode, how this table has to be received :  call by reference <-> call by value
		Dim newitab As Integer		'table was created by RfcGetData
	End Structure

	Public Const vTYPC As Integer = 0	 'RFC_CHAR       STRING $        characters

	' standard RFC functions
    Public Declare Function RfcOpenEx Lib "librfc32.dll" (ByVal Connect_param As String, ByVal error_info As RFC_ERROR_INFO_EX) As Integer
    Public Declare Function RfcOpenExt Lib "librfc32.dll" (ByVal dest As String, ByVal mode As RFC_MODE, ByVal Hostname As String, ByVal sysnr As Integer, ByVal gwhst As String, ByVal GwSrv As String, ByVal Client As String, ByVal user As String, ByVal pwd As String, ByVal lang As String, ByVal trace As Integer) As Integer
    Public Declare Sub RfcClose Lib "librfc32.dll" (ByVal Handle As Integer)
    Public Declare Function RfcLastError Lib "librfc32.dll" (ByVal error_info As RFC_ERROR_INFO) As Integer
	Public Declare Function RfcLastErrorEx Lib "librfc32.dll" (ByVal error_info As RFC_ERROR_INFO_EX) As Integer

	' Extended public RFC functions
	Public Declare Function RfcListen Lib "librfc32.dll" (ByVal hRFCServer As Integer) As Integer
	' Extended private RFC functions
	Declare Function RfcAllocParamSpace Lib "librfc32.dll" (ByVal numexp As Integer, ByVal numimp As Integer, ByVal numtab As Integer) As Integer
	Declare Function RfcFreeParamSpace Lib "librfc32.dll" (ByVal hParameterSpace As Integer) As Integer
	Declare Function RfcAddExportString Lib "librfc32.dll" Alias "RfcAddExportParam" (ByVal hParameterSpace As Integer, ByVal parpos As Integer, ByVal parname$, ByVal parnamelen As Integer, ByVal partype As Integer, ByVal parlen As Integer, ByVal par As String) As Integer
	Declare Function RfcAddImportParam Lib "librfc32.dll" (ByVal hParameterSpace As Integer, ByVal parpos As Integer, ByVal parname$, ByVal parnamelen As Integer, ByVal partype As Integer, ByVal parlen As Integer, ByVal par As String) As Integer
	Declare Function RfcDefineImportParam Lib "librfc32.dll" (ByVal hSpace As Integer, ByVal parpos As Integer, ByVal parname As String, ByVal parnamelen As Integer, ByVal partype As Integer, ByVal parlen As Integer) As Integer
	Declare Function RfcGetImportParam Lib "librfc32.dll" (ByVal hSpace As Integer, ByVal parpos As Integer, ByVal addr As String) As Integer
	Declare Function RfcAddTable Lib "librfc32.dll" (ByVal hParameterSpace As Integer, ByVal tabpos As Integer, ByVal tabname As String, ByVal tabnamelen As Integer, ByVal tabtype As Integer, ByVal tablen As Integer, ByVal tabhandle As Integer) As Integer
	Declare Function RfcCallExt Lib "librfc32.dll" (ByVal hRFCServer As Integer, ByVal hParameterSpace As Integer, ByVal funcname As String) As Integer
	Declare Function RfcReceiveExt Lib "librfc32.dll" (ByVal hRFCServer As Integer, ByVal hParameterSpace As Integer, ByVal exception As String) As Integer
	Declare Function RfcCallReceiveExt Lib "librfc32.dll" (ByVal hRFCServer As Integer, ByVal hParameterSpace As Integer, ByVal funcname As String, ByVal exception As String) As Integer
	Declare Function RfcAcceptExt Lib "librfc32.dll" (ByVal arguments As String) As Integer
	''Declare Function RfcDispatch Lib "librfc32.dll" (ByVal hRFCServer As Long, ByVal funcname$) As Long
	Declare Function RfcDispatch Lib "librfc32.dll" (ByVal hRFCServer As Integer) As Integer
	'Declare Function RfcGetNameEx Lib "librfc32.dll" (ByVal hRFCServer As Long, ByVal funcname$) As Long
	Declare Function RfcGetName Lib "librfc32.dll" (ByVal hRFCServer As Integer, ByVal funcname As String) As Integer
	Declare Function RfcGetDataExt Lib "librfc32.dll" (ByVal hRFCServer As Integer, ByVal hParameterSpace As Integer) As Integer
	Declare Function RfcGetData Lib "librfc32.dll" (ByVal hRFCServer As Integer, ByVal hParameters As Object, ByVal hTables As Object) As Integer
	Declare Function RfcGetTableHandle Lib "librfc32.dll" (ByVal hParameterSpace As Integer, ByVal tableno As Integer) As Integer
	Declare Function RfcSendDataExt Lib "librfc32.dll" (ByVal hRFCServer As Integer, ByVal hParameterSpace As Integer) As Integer
	Declare Function RfcRaise Lib "librfc32.dll" (ByVal hRFCServer As Integer, ByVal exception As String) As Integer
	'Declare Function RfcAbort Lib "librfc32.dll" (ByVal hRFCServer As Long, ByVal text$)
	'Declare Function RfcWaitForRequest Lib "librfc32.dll" (ByVal hRFCServer As Long, ByVal TimOut As Long) As Long

	' general table functions
	Declare Function ItCreate Lib "librfc32.dll" (ByVal ItName As String, ByVal ItRecLen As Integer, ByVal ItOccurs As Integer, ByVal mem As Integer) As Integer
	Declare Function ItDelete Lib "librfc32.dll" (ByVal hIt As Integer) As Integer
	Declare Function ItGetLine Lib "librfc32.dll" (ByVal hIt As Integer, ByVal ItLine As Integer) As Integer
	'Declare Function ItInsLine Lib "librfc32.dll" (ByVal hIt As Long, ByVal ItLine As Long) As Long
	Declare Function ItAppLine Lib "librfc32.dll" (ByVal hIt As Integer) As Integer
	Declare Function ItPutLine Lib "librfc32.dll" (ByVal hIt As Integer, ByVal lineNo As Integer, ByVal src As String) As Integer
	'Declare Function ItDelLine Lib "librfc32.dll" (ByVal hIt As Long, ByVal ItLine As Long) As Long
	Declare Function ItGupLine Lib "librfc32.dll" (ByVal hIt As Integer, ByVal ItLine As Integer) As Integer
	Declare Function ItCpyLine Lib "librfc32.dll" (ByVal hIt As Integer, ByVal ItLine As Integer, ByVal dest As String) As Integer
	'Declare Function ItFree Lib "librfc32.dll" (ByVal hIt As Long) As Long
	Declare Function ItFill Lib "librfc32.dll" (ByVal hIt As Integer) As Integer
	'Declare Function ItLeng Lib "librfc32.dll" (ByVal hIt As Long) As Long

	' kernel functions
	Declare Function lockWindowUpdate Lib "user32" Alias "lockWindowUpdateA" (ByVal hwnd As Short) As Integer
	'Declare Sub PointerToString Lib "MemSub.dll" Alias "MemCopy" (ByVal dest As String, ByVal Source As Long, ByVal nCount As Long)
	'Declare Sub StringToPointer Lib "MemSub.dll" Alias "MemCopy" (ByVal dest As Long, ByVal Source As String, ByVal nCount As Long)

	Public Declare Sub PointerToString Lib "kernel32" Alias "RtlCopyMemory" (ByVal dest As String, ByVal src As Integer, ByVal size As Integer)
	Public Declare Sub StringToPointer Lib "kernel32" Alias "RtlCopyMemory" (ByVal dest As Integer, ByVal src As String, ByVal size As Integer)

	'DELEGATED
	Public Delegate Function installDelegate(ByVal hRFCServer As Integer) As Integer
	Declare Function RfcInstallFunctionExt Lib "librfc32.dll" (ByVal hRFCServer As Integer, ByVal funcname As String, ByVal funcpointer As installDelegate, ByVal docu As String) As Integer


End Class