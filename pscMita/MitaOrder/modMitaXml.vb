Imports System.Xml
Module modMitaXml
	Public Sub xmlWriteChkPart(ByRef writer As XmlTextWriter, ByRef grp As Windows.Forms.GroupBox)
		Dim x As Windows.Forms.Control
		Dim c As Windows.Forms.CheckBox
		Dim index As Short = 0
		writer.WriteStartElement(grp.Name)
		For index = 0 To grp.Controls.Count - 1
			For Each x In grp.Controls
				If TypeOf x Is System.Windows.Forms.CheckBox Then
					c = x
					If index = c.TabIndex Then
						writer.WriteStartElement("Item")
						writer.WriteAttributeString("Text", CStr(c.Text))
						writer.WriteAttributeString("Index", CStr(c.TabIndex))
						writer.WriteAttributeString("Checked", CStr(c.Checked))
						writer.WriteAttributeString("Tag", CStr(c.Tag))
						writer.WriteAttributeString("Top", CStr(c.Top))
						writer.WriteAttributeString("Left", CStr(c.Left))
						writer.WriteEndElement()
					End If
				End If
			Next x
		Next index
		writer.WriteEndElement()
	End Sub

	Public Sub xmlWriteTypPart(ByRef writer As XmlTextWriter, ByRef grp As Windows.Forms.GroupBox)
		Dim x As Windows.Forms.Control
		Dim c As Windows.Forms.CheckBox
		writer.WriteStartElement(grp.Name)
		For Each x In grp.Controls
			If TypeOf x Is System.Windows.Forms.CheckBox Then
				c = x
				writer.WriteStartElement("Item")
				writer.WriteAttributeString("Text", CStr(c.Text))
				writer.WriteAttributeString("Checked", CStr(c.Checked))
				writer.WriteAttributeString("Tag", CStr(c.Tag))
				writer.WriteEndElement()
			End If
		Next x
		writer.WriteEndElement()
	End Sub
	Public Sub xmlWriteTxtPart(ByRef writer As XmlTextWriter, ByRef grp As Windows.Forms.GroupBox)
		Dim x As Windows.Forms.Control
		Dim c As Windows.Forms.TextBox
		Dim index As Short = 0
		writer.WriteStartElement(grp.Name)
		For Each x In grp.Controls
			If TypeOf x Is System.Windows.Forms.TextBox Then
				c = x
				If c.Text <> "" Then
					writer.WriteStartElement("Item")
					writer.WriteAttributeString("Tag", CStr(c.Tag))
					writer.WriteAttributeString("Text", CStr(c.Text))
					writer.WriteEndElement()
				End If
			End If
		Next x
		writer.WriteEndElement()
	End Sub
End Module
