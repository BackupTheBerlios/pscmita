Option Strict Off
Option Explicit On 
Imports VB = Microsoft.VisualBasic
Imports pscMitaDef.CMitaDef
Module modMitaOLP
	Public applicationForm As New frmMitaSRV
	Public Sub init()
		generateDLLs()
		startupInfo.id = 9
		startupInfo.typChar = "O_"
		startupInfo.formCaption = " - Batch Processor"
		startupInfo.iconIndex = 12
		startupInfo.workModus = workModi.inputDB

		setDefault("runType", "PROD")
		setDefault("alive", 10)
		setDefault("batchError", False)
		'setDefault("garbageRemove", True)
		setDefault("iconTray", False)
		setDefault("maxList", 250)
		setDefault("profile", False)
		setDefault("sameVersion", True)
		setDefault("sqlClass", mitaSqlClass.classError + mitaSqlClass.classOrder)
		setDefault("taskBar", True)
		setDefault("trace", "0")
		setDefault("hideMe", False)
		setDefault("miniSAP", False)
		setDefault("workModus", workModi.inputDB)
		setDefault("centerMe", False)
		setDefault("pool", "psc§sappool")
	End Sub
	Public Sub doProcess()
		colorProcess()
		sapToDB()
		colorRestore()
	End Sub
End Module