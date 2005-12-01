Option Strict Off
Option Explicit On 
Imports pscMitaDef.CMitaDef
Imports System.Type
Module MitaWriteDB

	Public Function addDBStructFieldEntry(ByVal Index As Short, ByRef sName As String, ByVal first As Integer, ByVal sLength As Integer, Optional ByVal removeOnly As Boolean = False) As Boolean
		Dim newVer As Integer
		Dim query As String
		Dim result As Boolean
		Dim hasOld As Boolean
		query = "SELECT version from " & mitaSystem.tableStructFields
		query = query & " WHERE sapstruct = '" & structureName & "'"
		query = query & " AND field = '" & sName & "'"
		query = query & " AND sapversionid = " & mitaSystem.sapSystemVERSIONID
		query = query & " ORDER BY version DESC"
		newVer = 0
		result = mitaConnect.queryExist(query, hasOld)
		If hasOld Then
			result = mitaConnect.queryNumber(query, newVer)
			newVer = newVer + 1
			If removeOnly Then newVer = newVer + 1
		Else
			newVer = 1
		End If
		If result And Not removeOnly Then
			query = "INSERT INTO " & mitaSystem.tableStructFields
			query = query & " (sapstruct, sapversionid, rfctype, field, version, first, recno, length, createid, createhost, createuser)"
			query = query & " VALUES("
			query = query & "'" & structureName & "'"
			query = query & ", " & mitaSystem.sapSystemVERSIONID
			query = query & ", " & mitaSystem.rfcType
			query = query & ", '" & sName & "'"
			query = query & ", " & newVer
			query = query & ", " & first
			query = query & ", " & Index
			query = query & ", " & sLength
			query = query & ", " & mitaData.createID
			query = query & ", '" & mitaData.createHost & "'"
			query = query & ", '" & mitaData.createUser & "'"
			query = query & ")"
			result = mitaConnect.execSQL(query)
		End If
		If result And hasOld Then
			result = inactivateOldStructField(structureName, sName, newVer)
		End If
		Return result
	End Function
	Public Function addDBCustFieldEntry(ByRef Index As Integer, ByRef field As FIELDSTRUCT, Optional ByVal removeOnly As Boolean = False) As Boolean
		Dim newVer As Integer
		Dim query As String
		Dim result As Boolean
		Dim hasOld As Boolean
		Dim name As String
		Dim first As Integer
		Dim structur As String
		Dim structField As String
		Dim length As Integer
		Dim typ As String
		Dim hash As Integer
		Dim level As Integer
		name = field.fName
		structField = field.fStructureField
		first = field.fFirst
		length = field.fLength
		typ = field.fType
		level = field.fLevel
		structur = field.fStructure
		query = "SELECT version from " & mitaSystem.tableCustFields
		query = query & " WHERE field = '" & name & "'"
		query = query & " AND sapsystemid = " & mitaSystem.sapSystemID
		query = query & " AND rfctype = " & mitaSystem.rfcType
		query = query & " ORDER BY version DESC"
		newVer = 0
		result = mitaConnect.queryExist(query, hasOld)
		If hasOld Then
			result = mitaConnect.queryNumber(query, newVer)
			newVer = newVer + 1
			If removeOnly Then newVer = newVer + 1
		Else
			newVer = 1
		End If
		If result And Not removeOnly Then
			query = "INSERT INTO " & mitaSystem.tableCustFields
			query = query & " (field, slevel, sapsystemid, rfctype, version,  sapstruct, sapfield, first, recno, length, ftype, createid, createhost, createuser)"
			query = query & " VALUES("
			query = query & "'" & name & "'"
			query = query & ", " & level
			query = query & ", " & mitaSystem.sapSystemID
			query = query & ", " & mitaSystem.rfcType
			query = query & ", " & newVer
			query = query & ", '" & structur & "'"
			query = query & ", '" & structField & "'"
			query = query & ", " & first
			query = query & ", " & Index
			query = query & ", " & length
			query = query & ", '" & typ & "'"
			query = query & ", " & mitaData.createID
			query = query & ", '" & mitaData.createHost & "'"
			query = query & ", '" & mitaData.createUser & "'"
			query = query & ")"
			result = mitaConnect.execSQL(query)
		End If
		If result And hasOld Then
			result = inactivateOldCustField(structur, name, newVer)
		End If
		Return result
	End Function

	Public Function addDbSqlEntry(ByRef sqlName As String, ByRef sqlQuery As String, ByRef sqlComment As String, Optional ByVal removeOnly As Boolean = False) As Boolean
		Dim newVer As Integer
		Dim query As String
		Dim result As Boolean
		Dim hasOld As Boolean
		addDbSqlEntry = False
		query = "SELECT version from " & mitaSystem.tableCustQuery
		query = query & " WHERE pscname = '" & UCase(sqlName) & "'"
		query = query & " AND sapsystemid = " & mitaSystem.sapSystemID
		query = query & " AND rfctype = " & mitaSystem.rfcType
		query = query & " ORDER BY version DESC"
		newVer = 0
		result = mitaConnect.queryExist(query, hasOld)
		If hasOld Then
			result = mitaConnect.queryNumber(query, newVer)
			newVer = newVer + 1
			If removeOnly Then newVer = newVer + 1
		Else
			newVer = 1
		End If
		If result And Not removeOnly Then
			query = "INSERT INTO " & mitaSystem.tableCustQuery
			query = query & " (pscname, sapsystemid, rfctype, version, createid, createhost, createuser, text, comments)"
			query = query & " VALUES("
			query = query & "'" & UCase(sqlName) & "'"
			query = query & ", " & mitaSystem.sapSystemID
			query = query & ", " & mitaSystem.rfcType
			query = query & ", " & newVer
			query = query & ", " & mitaData.createID
			query = query & ", '" & mitaData.createHost & "'"
			query = query & ", '" & mitaData.createUser & "'"
			query = query & ", '" & sqlQuery & "'"
			If sqlComment <> "" Then
				query = query & ", '" & sqlComment & "'"
			Else
				query = query & ", NULL"
			End If
			query = query & ")"
			result = mitaConnect.execSQL(query)
		End If
		If result And hasOld Then
			result = inactivateOldSQL(sqlName, newVer)
		End If
		Return result
	End Function

	Public Function addDbSapEntry(ByRef sSystem As String, ByRef Lins() As String, ByVal removeOnly As Boolean) As Boolean
		Dim newVer As Integer
		Dim query As String
		Dim result As Boolean
		Dim hasOld As Boolean
		Dim i As Short
		Dim j As Short
		Dim x() As String
		addDbSapEntry = False
		If Not IsNothing(Lins) Then
			For i = 0 To UBound(Lins)
				x = Split(Lins(i), "=", 2)
				Select Case Trim(UCase(x(0)))
					Case "ID"
						mitaSystem.sapSystemID = Val(x(1))
					Case "SAPGATEWAY"
						mitaSystem.sapSystemSAPGATEWAY = Trim(x(1))
					Case "SAPSERVICE"
						mitaSystem.sapSystemSAPSERVICE = Trim(x(1))
					Case "SAPID"
						mitaSystem.sapSystemSAPID = Trim(x(1))
					Case "SAPIDCLT"
						mitaSystem.sapSystemSAPIDCLT = Trim(x(1))
					Case "SAPUSER"
						mitaSystem.sapSystemSAPUSER = Trim(x(1))
					Case "SAPOWNER"
						mitaSystem.sapSystemSAPOWNER = Trim(x(1))
					Case "SAPSYSTEM"
						mitaSystem.sapSystemSAPSYSTEM = Trim(x(1))
					Case "SAPSERVER"
						mitaSystem.sapSystemSAPSERVER = Trim(x(1))
					Case "SAPCLIENT"
						mitaSystem.sapSystemSAPCLIENT = Trim(x(1))
					Case "DATABASE"
						mitaSystem.sapSystemDATABASE = Trim(x(1))
					Case "DATABASETYPE"
						mitaSystem.sapSystemDATABASETYPE = Trim(x(1))
					Case "DATABASEUSER"
						mitaSystem.sapSystemDATABASEUSER = Trim(x(1))
					Case "DATABASEPWD"
						mitaSystem.sapSystemDATABASEPWD = Trim(x(1))
					Case "SAPSYSTEMVERSIONID"
						mitaSystem.sapSystemVERSIONID = Val(x(1))
					Case "ENVIRON"
						mitaSystem.runType = x(1)
				End Select
			Next i
		End If
		query = "SELECT version from " & mitaSystem.tableSapSystems
		query = query & " WHERE sapname = '" & sSystem & "'"
		query = query & " ORDER BY version DESC"
		newVer = 0
		result = mitaConnect.queryExist(query, hasOld)
		If hasOld Then
			result = mitaConnect.queryNumber(query, newVer)
			newVer = newVer + 1
			If removeOnly Then newVer = newVer + 1
		Else
			newVer = 1
		End If
		If result And Not removeOnly Then
			query = "INSERT INTO " & mitaSystem.tableSapSystems
			query = query & " (sapsystemid"
			query = query & ", sapname"
			query = query & ", version"
			query = query & ", sapversionid"
			query = query & ", SAPGATEWAY"
			query = query & ", SAPSERVICE"
			query = query & ", SAPID"
			query = query & ", SAPIDCLT"
			query = query & ", SAPUSER"
			query = query & ", SAPOWNER"
			query = query & ", SAPSYSTEM"
			query = query & ", SAPSERVER"
			query = query & ", SAPCLIENT"
			query = query & ", databas"
			query = query & ", databastype"
			query = query & ", databasuser"
			query = query & ", databaspwd"
			query = query & ", createUser"
			query = query & ", createID"
			query = query & ", createHost"
			query = query & ", environment)"

			query = query & " VALUES("
			query = query & mitaSystem.sapSystemID
			query = query & ", '" & sSystem & "'"
			query = query & ", " & newVer
			query = query & ", " & mitaSystem.sapSystemVERSIONID
			query = query & ", '" & mitaSystem.sapSystemSAPGATEWAY & "'"
			query = query & ", '" & mitaSystem.sapSystemSAPSERVICE & "'"
			query = query & ", '" & mitaSystem.sapSystemSAPID & "'"
			query = query & ", '" & mitaSystem.sapSystemSAPIDCLT & "'"
			query = query & ", '" & mitaSystem.sapSystemSAPUSER & "'"
			query = query & ", '" & mitaShared.DoXor(mitaSystem.sapSystemSAPOWNER) & "'"
			query = query & ", '" & mitaSystem.sapSystemSAPSYSTEM & "'"
			query = query & ", '" & mitaSystem.sapSystemSAPSERVER & "'"
			query = query & ", '" & mitaSystem.sapSystemSAPCLIENT & "'"
			query = query & ", '" & mitaSystem.sapSystemDATABASE & "'"
			query = query & ", '" & mitaSystem.sapSystemDATABASETYPE & "'"
			query = query & ", '" & mitaSystem.sapSystemDATABASEUSER & "'"
			query = query & ", '" & mitaShared.DoXor(mitaSystem.sapSystemDATABASEPWD) & "'"
			query = query & ", '" & mitaData.createUser & "'"
			query = query & ", " & mitaData.createID
			query = query & ", '" & mitaData.createHost & "'"
			query = query & ", '" & mitaSystem.runType & "'"
			query = query & ")"
			result = mitaConnect.execSQL(query)
		End If
		If result And hasOld Then
			result = inactivateOldSap(sSystem, newVer)
		End If
		Return result
	End Function

	Public Function transferDbSapEntry(ByRef sName As String, ByVal fromConnect As String, ByVal toConnect As String) As Boolean
		Dim newVer As Integer
		Dim query As String
		Dim result As Boolean
		Dim hasOld As Boolean
		Dim sav As String = mitaSystem.sapSystemNAME
		mitaConnect.dbConnectString = fromConnect
		mitaShared.readDBSapSystemFromName(sName)
		newVer = 1
		mitaConnect.dbConnectString = toConnect
		query = "SELECT version from " & mitaSystem.tableSapSystems
		query = query & " WHERE sapname = '" & sName & "'"
		query = query & " ORDER BY version DESC"
		result = mitaConnect.queryExist(query, hasOld)
		If hasOld Then Return True
		query = "INSERT INTO " & mitaSystem.tableSapSystems
		query = query & " (sapsystemid"
		query = query & ", sapname"
		query = query & ", version"
		query = query & ", sapversionid"
		query = query & ", SAPGATEWAY"
		query = query & ", SAPSERVICE"
		query = query & ", SAPID"
		query = query & ", SAPIDCLT"
		query = query & ", SAPUSER"
		query = query & ", SAPOWNER"
		query = query & ", SAPSYSTEM"
		query = query & ", SAPSERVER"
		query = query & ", SAPCLIENT"
		query = query & ", databas"
		query = query & ", databastype"
		query = query & ", databasuser"
		query = query & ", databaspwd"
		query = query & ", createUser"
		query = query & ", createID"
		query = query & ", createHost"
		query = query & ", environment)"

		query = query & " VALUES("
		query = query & mitaSystem.sapSystemId
		query = query & ", '" & sName & "'"
		query = query & ", " & newVer
		query = query & ", " & mitaSystem.sapSystemVERSIONID
		query = query & ", '" & mitaSystem.sapSystemSAPGATEWAY & "'"
		query = query & ", '" & mitaSystem.sapSystemSAPSERVICE & "'"
		query = query & ", '" & mitaSystem.sapSystemSAPID & "'"
		query = query & ", '" & mitaSystem.sapSystemSAPIDCLT & "'"
		query = query & ", '" & mitaSystem.sapSystemSAPUSER & "'"
		query = query & ", '" & mitaShared.DoXor(mitaSystem.sapSystemSAPOWNER) & "'"
		query = query & ", '" & mitaSystem.sapSystemSAPSYSTEM & "'"
		query = query & ", '" & mitaSystem.sapSystemSAPSERVER & "'"
		query = query & ", '" & mitaSystem.sapSystemSAPCLIENT & "'"
		query = query & ", '" & mitaSystem.sapSystemDATABASE & "'"
		query = query & ", '" & mitaSystem.sapSystemDATABASETYPE & "'"
		query = query & ", '" & mitaSystem.sapSystemDATABASEUSER & "'"
		query = query & ", '" & mitaShared.DoXor(mitaSystem.sapSystemDATABASEPWD) & "'"
		query = query & ", '" & mitaData.createUser & "'"
		query = query & ", " & mitaData.createID
		query = query & ", '" & mitaData.createHost & "'"
		query = query & ", '" & mitaSystem.runType & "'"
		query = query & ")"
		result = mitaConnect.execSQL(query)
		mitaShared.readDBSapSystemFromName(sav)
		mitaConnect.dbConnectString = mitaSystem.connectString
		Return result
	End Function

	Public Function deleteDbSapEntry(ByRef sName As String, ByVal fromConnect As String) As Boolean
		Dim query As String
		Dim result As Boolean
		mitaConnect.dbConnectString = fromConnect
		query = "UPDATE " & mitaSystem.tableSapSystems
		query = query & " SET deleted = 'Y'"
		query = query & " WHERE sapname = '" & sName & "'"
		result = mitaConnect.execSQL(query)
		Return result
	End Function

	Public Function addDbSapVersionEntry(ByRef funct As String, ByRef struct As String, ByRef versionId As Integer, ByRef versionName As String, ByRef subVersion As String, Optional ByVal removeOnly As Boolean = False) As Boolean
		Dim newVer As Integer
		Dim query As String
		Dim result As Boolean
		Dim hasOld As Boolean
		Dim i As Short
		Dim j As Short
		Dim x() As String
		query = "SELECT version from " & mitaSystem.tableSapVersions
		query = query & " WHERE sapversionid = " & versionId
		query = query & " ORDER BY version DESC"
		newVer = 0
		result = mitaConnect.queryExist(query, hasOld)
		If hasOld Then
			result = mitaConnect.queryNumber(query, newVer)
			newVer = newVer + 1
			If removeOnly Then newVer = newVer + 1
		Else
			newVer = 1
		End If
		If result And Not removeOnly Then
			query = "INSERT INTO " & mitaSystem.tableSapVersions
			query = query & " (sapversionid"
			query = query & ", sapversionname"
			query = query & ", sapsubversion"
			query = query & ", version"
			query = query & ", createUser"
			query = query & ", createID"
			query = query & ", replacestruct"
			query = query & ", replacefunc"
			query = query & ", createHost)"

			query = query & " VALUES("
			query = query & versionId
			query = query & ", '" & versionName & "'"
			query = query & ", '" & subVersion & "'"
			query = query & ", " & newVer
			query = query & ", '" & mitaData.createUser & "'"
			query = query & ", " & mitaData.createID
			query = query & ", '" & struct & "'"
			query = query & ", '" & funct & "'"
			query = query & ", '" & mitaData.createHost & "'"
			query = query & ")"
			result = mitaConnect.execSQL(query)
		End If
		If result And hasOld Then
			result = inactivateOldSapVersion(versionId, newVer)
		End If
		Return result
	End Function

	Private Function inactivateOldSQL(ByRef pscName As String, ByRef curVer As Integer) As Boolean
		Dim query As String
		query = "UPDATE " & mitaSystem.tableCustQuery
		query = query & " SET activ = 'N'"
		query = query & " WHERE pscname = '" & UCase(pscName) & "'"
		query = query & " AND sapsystemid = " & mitaSystem.sapSystemID
		query = query & " AND rfctype = " & mitaSystem.rfcType
		query = query & " AND version < "
		query = query & curVer
		Return mitaConnect.execSQL(query)
	End Function
	Private Function inactivateOldSap(ByRef sapName As String, ByRef curVer As Integer) As Boolean
		Dim query As String
		query = "UPDATE " & mitaSystem.tableSapSystems
		query = query & " SET activ = 'N'"
		query = query & " WHERE sapname = '" & sapName & "'"
		query = query & " AND version < "
		query = query & curVer
		Return mitaConnect.execSQL(query)
	End Function

	Private Function inactivateOldSapVersion(ByRef sapversionid As String, ByRef curVer As Integer) As Boolean
		Dim query As String
		query = "UPDATE " & mitaSystem.tableSapVersions
		query = query & " SET activ = 'N'"
		query = query & " WHERE sapversionid = " & sapversionid
		query = query & " AND version < "
		query = query & curVer
		Return mitaConnect.execSQL(query)
	End Function
	Public Function addDbStructEntry(ByRef structur As String, ByRef Index As Integer, ByRef length As Integer, ByVal rfcFunc As String, ByVal level As Integer, ByVal target As String, ByVal typ As String, Optional ByVal removeOnly As Boolean = False) As Boolean
		Dim newVer As Integer
		Dim query As String
		Dim result As Boolean = True
		Dim hasOld As Boolean
		If result Then
			query = "SELECT version from " & mitaSystem.tableStructures
			query = query & " WHERE sapstruct = '" & structur & "'"
			query = query & " AND sapversionid = " & mitaSystem.sapSystemVERSIONID
			query = query & " AND rfctype = " & mitaSystem.rfcType
			query = query & " ORDER BY version DESC"
			newVer = 0
			result = result And mitaConnect.queryExist(query, hasOld)
			If hasOld Then
				result = result And mitaConnect.queryNumber(query, newVer)
				newVer = newVer + 1
				If removeOnly Then newVer = newVer + 1
			Else
				newVer = 1
			End If
			If result And Not removeOnly Then
				query = "INSERT INTO " & mitaSystem.tableStructures
				query = query & " (sapstruct, slevel, sapversionid, rfctype, rfcfunction, version, recno, length,  createid, createhost, createuser, datatype, datastruct)"
				query = query & " VALUES("
				query = query & "'" & structur & "'"
				query = query & ", " & level
				query = query & ", " & mitaSystem.sapSystemVERSIONID
				query = query & ", " & mitaSystem.rfcType
				query = query & ", '" & rfcFunc & "'"
				query = query & ", " & newVer
				query = query & ", " & Index
				query = query & ", " & length
				query = query & ", " & mitaData.createID
				query = query & ", '" & mitaData.createHost & "'"
				query = query & ", '" & mitaData.createUser & "'"
				query = query & ", '" & typ & "'"
				query = query & ", '" & target & "'"
				query = query & ")"
				result = result And mitaConnect.execSQL(query)
			End If
			If result And hasOld Then
				result = result And inactivateOldStruct(structur, newVer)
			End If
		End If
		Return result
	End Function

	Public Function addDbEventEntry(ByRef eventName As String, ByRef runType As String, ByRef eventAction As String, ByRef eventParameter As String, ByRef eventComment As String, ByVal Index As Integer, Optional ByVal removeOnly As Boolean = False) As Boolean
		Dim newVer As Integer
		Dim x() As String
		Dim query As String
		Dim result As Boolean
		Dim hasOld As Boolean
		addDbEventEntry = False
		query = "SELECT version from " & mitaSystem.tableEventControl
		query = query & " WHERE event = '" & eventName & "'"
		query = query & " AND sapsystemid = " & mitaSystem.sapSystemID
		query = query & " AND rfctype = " & mitaSystem.rfcType
		query = query & " AND runtype = '" & runType & "'"
		query = query & " AND activ = 'Y'"
		query = query & " AND action IS NULL"
		query = query & " ORDER BY version DESC"
		newVer = 0
		result = mitaConnect.queryExist(query, hasOld)
		If hasOld Then
			result = mitaConnect.queryNumber(query, newVer)
			newVer = newVer + 1
			If removeOnly Then newVer = newVer + 1
		Else
			newVer = 1
		End If
		If result And Not removeOnly Then
			If Index = -1 Then
				query = "INSERT INTO " & mitaSystem.tableEventControl
				query = query & " (event, sapsystemid, rfctype, runtype, version, masterversion, recno, createid, createhost, createuser, comments)"
				query = query & " VALUES("
				query = query & "'" & eventName & "'"
				query = query & ", " & mitaSystem.sapSystemID
				query = query & ", " & mitaSystem.rfcType
				query = query & ", '" & runType & "'"
				query = query & ", " & newVer
				query = query & ", " & newVer
				query = query & ", " & Index
				query = query & ", " & mitaData.createID
				query = query & ", '" & mitaData.createHost & "'"
				query = query & ", '" & mitaData.createUser & "'"
				query = query & ", '" & eventComment & "'"
				query = query & ")"
			Else
				query = "INSERT INTO " & mitaSystem.tableEventControl
				query = query & " (event, sapsystemid, rfctype, runtype, version, masterversion, recno, action, parameter, createid, createhost, createuser, comments)"
				query = query & " VALUES("
				query = query & "'" & eventName & "'"
				query = query & ", " & mitaSystem.sapSystemID
				query = query & ", " & mitaSystem.rfcType
				query = query & ", '" & runType & "'"
				query = query & ", " & newVer
				query = query & ", " & newVer
				query = query & ", " & Index
				query = query & ", '" & eventAction & "'"
				query = query & ", '" & Replace(eventParameter, "'", "~") & "'"
				query = query & ", " & mitaData.createID
				query = query & ", '" & mitaData.createHost & "'"
				query = query & ", '" & mitaData.createUser & "'"
				query = query & ", '" & eventComment & "'"
				query = query & ")"
			End If
			result = mitaConnect.execSQL(query)
		End If
		If result And Index = -1 Then
			query = "UPDATE " & mitaSystem.tableEventControl
			query = query & " SET masterversion = " & newVer
			query = query & " WHERE masterversion = 0"
			result = mitaConnect.execSQL(query)
		End If
		If result And hasOld Then
			result = inactivateOldEvent(eventName, eventAction, newVer, runType)
		End If
		Return result
	End Function


	Private Function inactivateOldCustField(ByRef structur As String, ByRef field As String, ByRef curVer As Integer) As Boolean
		Dim query As String
		query = "UPDATE " & mitaSystem.tableCustFields
		query = query & " SET activ = 'N'"
		query = query & " WHERE field = '" & field & "'"
		query = query & " AND sapsystemid = " & mitaSystem.sapSystemID
		query = query & " AND rfctype = " & mitaSystem.rfcType
		query = query & " AND version < " & curVer
		Return mitaConnect.execSQL(query)
	End Function
	Public Function addDbTableEntry(ByRef name As String, ByRef tLeft As String, ByRef tRight As String, Optional ByVal removeOnly As Boolean = False) As Boolean
		Dim newVer As Integer
		Dim query As String
		Dim result As Boolean
		Dim hasOld As Boolean
		addDbTableEntry = False
		query = "SELECT version FROM " & mitaSystem.tableCustTables
		query = query & " WHERE tablename = '" & name & "'"
		query = query & " AND sapsystemid = " & mitaSystem.sapSystemID
		query = query & " AND tleft = '" & tLeft & "'"
		query = query & " ORDER BY version DESC"
		newVer = 0
		result = mitaConnect.queryExist(query, hasOld)
		If hasOld Then
			result = mitaConnect.queryNumber(query, newVer)
			newVer = newVer + 1
			If removeOnly Then newVer = newVer + 1
		Else
			newVer = 1
		End If
		If result And Not removeOnly Then
			query = "INSERT INTO " & mitaSystem.tableCustTables
			query = query & " (tablename, sapsystemid, rfctype, version, createid, createhost, createuser, tleft, tright)"
			query = query & " VALUES("
			query = query & "'" & name & "'"
			query = query & ", " & mitaSystem.sapSystemID
			query = query & ", " & mitaSystem.rfcType
			query = query & ", " & newVer
			query = query & ", " & mitaData.createID
			query = query & ", '" & mitaData.createHost & "'"
			query = query & ", '" & mitaData.createUser & "'"
			If tLeft = "" Then
				query = query & ", NULL"
			Else
				query = query & ", '" & tLeft & "'"
			End If
			If tRight = "" Then
				query = query & ", NULL"
			Else
				query = query & ", '" & tRight & "'"
			End If
			query = query & ")"
			result = mitaConnect.execSQL(query)
		End If
		If result And hasOld Then
			result = inactivateOldTables(name, tLeft, newVer)
		End If
		Return result
	End Function
	Public Function addDbCombiEntry(ByRef name As String, ByRef items As String, Optional ByVal removeOnly As Boolean = False) As Boolean
		Dim newVer As Integer
		Dim query As String
		Dim result As Boolean
		Dim hasOld As Boolean
		addDbCombiEntry = False
		query = "SELECT version FROM " & mitaSystem.tableCustCombis
		query = query & " WHERE combiname = '" & name & "'"
		query = query & " ORDER BY version DESC"
		newVer = 0
		result = mitaConnect.queryExist(query, hasOld)
		If hasOld Then
			result = mitaConnect.queryNumber(query, newVer)
			newVer = newVer + 1
			If removeOnly Then newVer = newVer + 1
		Else
			newVer = 1
		End If
		If result And Not removeOnly Then
			query = "INSERT INTO " & mitaSystem.tableCustCombis
			query = query & " (combiname, version, createid, createhost, createuser, items)"
			query = query & " VALUES("
			query = query & "'" & name & "'"
			query = query & ", " & newVer
			query = query & ", " & mitaData.createID
			query = query & ", '" & mitaData.createHost & "'"
			query = query & ", '" & mitaData.createUser & "'"
			query = query & ", '" & items & "'"
			query = query & ")"
			result = mitaConnect.execSQL(query)
		End If
		If result And hasOld Then
			result = inactivateOldCombis(name, newVer)
		End If
		Return result
	End Function

	Private Function inactivateOldTables(ByRef tableName As String, ByRef tLeft As String, ByRef curVer As Integer) As Boolean
		Dim query As String
		query = "UPDATE " & mitaSystem.tableCustTables
		query = query & " SET activ = 'N'"
		query = query & " WHERE tablename = '" & tableName & "'"
		query = query & " AND tleft = '" & tLeft & "'"
		query = query & " AND sapsystemid = " & mitaSystem.sapSystemID
		query = query & " AND version < "
		query = query & curVer
		Return mitaConnect.execSQL(query)
	End Function


	Private Function inactivateOldCombis(ByRef combiName As String, ByRef curVer As Integer) As Boolean
		Dim query As String
		query = "UPDATE " & mitaSystem.tableCustCombis
		query = query & " SET activ = 'N'"
		query = query & " WHERE combiname = '" & combiName & "'"
		query = query & " AND version < "
		query = query & curVer
		Return mitaConnect.execSQL(query)
	End Function

	Private Function inactivateOldStruct(ByRef structur As String, ByRef curVer As Integer) As Boolean
		Dim query As String
		query = "UPDATE " & mitaSystem.tableStructures
		query = query & " SET activ = 'N'"
		query = query & " WHERE sapstruct = '" & structur & "'"
		query = query & " AND sapversionid = " & mitaSystem.sapSystemVERSIONID
		query = query & " AND rfctype = " & mitaSystem.rfcType
		query = query & " AND version < " & curVer
		Return mitaConnect.execSQL(query)
	End Function
	Private Function inactivateOldStructField(ByRef structur As String, ByRef field As String, ByRef curVer As Integer) As Boolean
		Dim query As String
		query = "UPDATE " & mitaSystem.tableStructFields
		query = query & " SET activ = 'N'"
		query = query & " WHERE sapstruct = '" & structur & "'"
		query = query & " AND field = '" & field & "'"
		query = query & " AND sapversionid = " & mitaSystem.sapSystemVERSIONID
		query = query & " AND rfctype = " & mitaSystem.rfcType
		query = query & " AND version < "
		query = query & curVer
		Return mitaConnect.execSQL(query)
	End Function
	Private Function inactivateOldEvent(ByRef eventNam As String, ByRef action As String, ByRef curVer As Integer, ByRef rt As String) As Boolean
		Dim query As String
		query = "UPDATE " & mitaSystem.tableEventControl
		query = query & " SET activ = 'N'"
		query = query & " WHERE event = '" & eventNam & "'"
		If action <> "" Then
			query = query & " AND action = '" & action & "'"
		Else
			query = query & " AND masterversion < "
			query = query & curVer
		End If
		query = query & " AND sapsystemid = " & mitaSystem.sapSystemID
		query = query & " AND rfctype = " & mitaSystem.rfcType
		query = query & " AND runtype = '" & rt & "'"
		query = query & " AND version < "
		query = query & curVer
		Return mitaConnect.execSQL(query)
	End Function
	Public Function copyContent(ByVal tableName As String, ByVal oldId As Integer, ByVal newId As Integer, ByVal versionID As Integer)
		Dim query As String
		Dim insert As String
		Dim inserts() As String
		Dim insertCount As Integer = -1
		Dim idbc As OdbcCommand = mitaConnect.odbc_connection.CreateCommand()
		Dim reader As OdbcDataReader
		Dim i As Integer
		Dim f$
		query = "SELECT * from " & tableName
		query = query & " WHERE activ = 'Y'"
		query = query & " AND sapsystemid = " & mitaSystem.sapSystemID
		idbc.CommandText = query
		On Error GoTo isErr
		mitaConnect.odbc_connection.Open()
		reader = idbc.ExecuteReader()
		While reader.Read
			insert = "INSERT into " & tableName & " VALUES("
			For i = 0 To reader.FieldCount - 1
				f$ = reader.GetName(i)
				If i > 0 Then insert = insert & ", "
				Select Case f$
					Case "sapsystemversionid"
						insert = insert & CStr(versionID)
					Case "SAPSYSTEMID"
						insert = insert & CStr(newId)
					Case "CREATETIME"
						insert = insert & " sysdate"
					Case "CREATEUSER"
						insert = insert & " '" & mitaData.createUser & "'"
					Case "CREATEHOST"
						insert = insert & " '" & mitaData.createHost & "'"
					Case "CREATEID"
						insert = insert & " " & mitaData.createID
					Case Else
						If TypeOf reader.Item(i) Is String Then
							insert = insert & "'"
						End If
						insert = insert & CStr(reader.GetValue(i))
						If TypeOf reader.Item(i) Is String Then
							insert = insert & "'"
						End If
				End Select
			Next i
			insert = insert & ")"
			insertCount = insertCount + 1
			ReDim Preserve inserts(insertCount)
			inserts(insertCount) = insert
		End While
		mitaConnect.odbc_connection.Close()
		copyContent = True
		For i = 0 To insertCount
			copyContent = copyContent And mitaConnect.execSQL(inserts(i))
		Next
Exx:
		If Not IsNothing(reader) Then reader.Close()
		If Not IsNothing(idbc) Then idbc.Dispose()
		If mitaConnect.odbc_connection.State = ConnectionState.Open Then mitaConnect.odbc_connection.Close()
		Exit Function
isErr:
		mitaData.errorDescription = Err.Description & vbCrLf & query & vbCrLf & insert
#If Not MitaOrder Then
		'saporder.eventRaise(errorDescription, mitaEventCodes.errorDataBase, "readDbCustCombis")
#End If
		Resume Exx

	End Function
End Module