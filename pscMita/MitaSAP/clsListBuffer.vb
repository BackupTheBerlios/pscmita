Public Class ListBuffer
	Private maxList As Integer
	Private listCounter As Integer = 0
	Private listSize As Integer = 0
	Dim myList() As String
	Public Sub setMaxList(ByVal maxSize As Integer)
		maxList = maxSize - 1
		ReDim myList(maxList)
	End Sub
	Public Sub addItem(ByVal item As String)
		myList(listCounter) = item
		If listSize <= maxList Then listSize = listSize + 1
		listCounter = listCounter + 1
		If listCounter > maxList Then
			listCounter = 0
		End If
	End Sub
	Public Function readItems() As String()
		Dim tmp() As String
		Dim i As Integer
		Dim j As Integer
		If listSize = 0 Then Return Nothing
		If listSize < maxList Then
			ReDim tmp(listSize - 1)
			For i = 0 To listCounter - 1
				tmp(i) = myList(i)
			Next
		Else
			ReDim tmp(maxList)
			j = 0
			For i = listCounter To maxList
				tmp(j) = myList(i)
				j = j + 1
			Next
			For i = 0 To listCounter - 1
				tmp(j) = myList(i)
				j = j + 1
			Next
		End If
		Return tmp
	End Function

End Class
