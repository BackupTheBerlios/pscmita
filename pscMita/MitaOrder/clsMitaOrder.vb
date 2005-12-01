Option Strict Off
Option Explicit On 
Imports VB = Microsoft.VisualBasic
Imports System.Data.Odbc
Imports pscSapServer.CSapServer
Imports pscMitaDef.CMitaDef
'<ComClass(MitaOrder.ClassId, MitaOrder.InterfaceId, MitaOrder.EventsId)> 
Public Class CMitaOrder

	'#Region "COM GUIDs"
	'	' These  GUIDs provide the COM identity for this class 
	'	' and its COM interfaces. If you change them, existing 
	'	' clients will no longer be able to access the class.
	'	Public Const ClassId As String = "35AC9201-15C4-4d72-9CB5-3BC9F1B18B37"
	'	Public Const InterfaceId As String = "B0645340-C106-4aef-B343-BBEFD71366BF"
	'	Public Const EventsId As String = "33A82500-04C1-47d4-9775-ACD5605D28DB"
	'#End Region

	Private Const hashCodeCount As Integer = 1023

	Private Const cUpdateControl As String = "ÿÿ