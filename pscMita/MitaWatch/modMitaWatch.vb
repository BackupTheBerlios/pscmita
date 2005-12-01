Module modMitaWatch
	Friend mitaData As New pscMitaData.CMitaData
	Friend mitaShared As New pscMitaShared.CMitaShared
	Friend mitaConnect As New pscMitaConnect.CMitaConnect
	Friend mitaSystem As New pscMitaSapSystem.CMitaSapSystem
	Friend mitaApplication As String
	Public mainFrm As frmMitaWatch
	Public Sub enableButtons(ByVal dummy1 As System.Windows.Forms.Form, ByVal dummy2 As Integer)
		'Dummy 
	End Sub
	Public Sub disableButtons(ByVal dummy1 As System.Windows.Forms.Form)
		'Dummy
	End Sub
End Module
