Public Class CSapPool
	Implements IDisposable
		Public spsAdNo As Integer = -1
		Public spsMotivno As Integer = -1
		Public spsComboNo As Integer = -1
		Public spsPosNo As Integer = -1
		Public spsYSize As Single = 0
		Public spsXSize As Single = 0
		Public spsYLoc As Single = 0
		Public spsXLoc As Single = 0
		Public spsAdType As String = ""
		Public spsAvm As String = ""
		Public spsVno As String = ""
		Public spsDsnNam As String = ""
		Public spsTxtNam As String = ""
		Public spsBoxNo As String = ""
		Public spsAdSort As String = ""
		Public spsClientNo As String = "000"

		Public Overloads Sub Dispose() Implements IDisposable.Dispose
			Dispose(True)
			GC.SuppressFinalize(Me)
		End Sub

		Protected Overridable Overloads Sub Dispose(ByVal disposing As Boolean)
			If disposing Then
				' Free other state (managed objects).
			End If
			spsAdType = Nothing
			spsClientNo = Nothing
			spsAvm = Nothing
			spsVno = Nothing
			spsDsnNam = Nothing
			spsTxtNam = Nothing
			spsBoxNo = Nothing
			spsAdSort = Nothing
		End Sub

		Protected Overrides Sub Finalize()
			Dispose(False)
		End Sub
	Public Overloads Function equals(ByVal cmp As CSapPool)
		If spsAdNo = cmp.spsAdNo _
		And spsVno = cmp.spsVno _
		And spsMotivno = cmp.spsMotivno _
		And spsComboNo = cmp.spsComboNo _
		And spsAvm = cmp.spsAvm _
		And spsDsnNam = cmp.spsDsnNam _
		And spsTxtNam = cmp.spsTxtNam _
		And spsBoxNo = cmp.spsBoxNo _
		And spsAdSort = cmp.spsAdSort _
		And spsYSize = cmp.spsYSize _
		And spsXSize = cmp.spsXSize _
		And spsYLoc = cmp.spsYLoc _
		And spsXLoc = cmp.spsXLoc _
		And spsPosNo = cmp.spsPosNo _
		And spsAdType = cmp.spsAdType _
		And spsClientNo = cmp.spsClientNo _
		Then
			Return True
		Else
			Return False
		End If
	End Function
	Public Function clone() As CSapPool
		Dim cnl As New CSapPool
		cnl.spsAdNo = spsAdNo
		cnl.spsVno = spsVno
		cnl.spsComboNo = spsComboNo
		cnl.spsPosNo = spsPosNo
		cnl.spsYSize = spsYSize
		cnl.spsXSize = spsXSize
		cnl.spsYLoc = spsYLoc
		cnl.spsXLoc = spsXLoc
		cnl.spsMotivno = spsMotivno
		cnl.spsAvm = spsAvm
		cnl.spsDsnNam = spsDsnNam
		cnl.spsTxtNam = spsTxtNam
		cnl.spsBoxNo = spsBoxNo
		cnl.spsAdSort = spsAdSort
		cnl.spsAdType = spsAdType
		cnl.spsClientNo = spsClientNo
		Return cnl
	End Function
End Class
