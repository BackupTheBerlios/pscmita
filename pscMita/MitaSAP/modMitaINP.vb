Imports VB = Microsoft.VisualBasic
Imports pscMitaDef.CMitaDef
Module modMitaINP
	Public applicationForm As New frmMitaSRV
	Public Sub init()
		generateDLLs()
		startupInfo.runType = ""
		startupInfo.id = 19
		startupInfo.typChar = "I_"
		startupInfo.formCaption = " - Input Server"
		startupInfo.workModus = workModi.inputSap
		startupInfo.iconIndex = 11

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
		'setDefault("pool", "pscsappool")

		SapOrder.optionsSaveAllSAP(True)
	End Sub
	Public Sub doProcess()
		Dim dummy As Long
		colorDB()
		SapOrder.orderWriteDB(dummy)
		colorRestore()
	End Sub
End Module
