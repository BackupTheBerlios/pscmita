Imports pscMitaDef.CMitaDef
Module modMitaStartup
	Public Enum workModi
		inputDirectory
		inputSap
		inputDB
	End Enum
	Public Structure defaultStructure
		Dim workModus As workModi
		Dim sapSystemName As String
		Dim runType As String
		Dim alive As Integer
		Dim garbageRemove As Boolean
		Dim sameVersion As Boolean
		Dim sqlClass As mitaSqlClass
		Dim miniSAP As Boolean
		Dim maxList As Integer
		Dim trace As Char
		Dim pool As String
		Dim profile As Boolean
		Dim taskBar As Boolean
		Dim iconTray As Boolean
		Dim hideMe As Boolean
		Dim centerMe As Boolean
		Dim batchError As Boolean
	End Structure
	Public Structure usedStructure
		Dim workModus As Boolean
		Dim runType As Boolean
		Dim alive As Boolean
		Dim garbageRemove As Boolean
		Dim sameVersion As Boolean
		Dim sqlClass As Boolean
		Dim miniSAP As Boolean
		Dim maxList As Boolean
		Dim trace As Boolean
		Dim pool As Boolean
		Dim profile As Boolean
		Dim taskBar As Boolean
		Dim iconTray As Boolean
		Dim hideMe As Boolean
		Dim centerMe As Boolean
		Dim batchError As Boolean
	End Structure

	Public defaultInfo As defaultStructure
	Public usedInfo As usedStructure
End Module
