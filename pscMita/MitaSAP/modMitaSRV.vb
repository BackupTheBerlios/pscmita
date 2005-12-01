Option Strict Off
Option Explicit On 
Imports VB = Microsoft.VisualBasic
Imports pscMitaDef.CMitaDef
Module modMitaSRV
	Public applicationForm As frmMitaSRV
	Public Sub init()
		generateDLLs()
		startupInfo.runType = ""
		startupInfo.id = 29
		startupInfo.typChar = "S_"
		startupInfo.formCaption = " - Input Processor"
		startupInfo.iconIndex = 13
		startupInfo.workModus = workModi.inputSap

		setDefault("runType", "PROD")
		setDefault("alive", 10)
		setDefault("batchError", False)
		setDefault("garbageRemove", True)
		setDefault("iconTray", False)
		setDefault("maxList", 100)
		setDefault("profile", False)
		setDefault("sameVersion", True)
		setDefault("sqlClass", mitaSqlClass.classError + mitaSqlClass.classOrder)
		setDefault("taskBar", True)
		setDefault("trace", "0")
		setDefault("hideMe", False)
		setDefault("miniSAP", False)
		setDefault("workModus", workModi.inputSap)
		setDefault("centerMe", False)
		setDefault("pool", "pscsappool")
	End Sub

	Public Sub doProcess()
		colorDB()
		If SapOrder.orderWriteDB(Nothing) Then
			colorProcess()
			sapToDB()
			colorRestore()
		Else
			' some errors do not prohibit to process!
			Dim code As mitaErrorCodes = SapOrder.orderGetErrorCode
			If code And mitaErrorCodes.orderOK > 0 Then
				colorProcess()
				sapToDB()
				colorRestore()
			End If
		End If
		colorRestore()
	End Sub
End Module