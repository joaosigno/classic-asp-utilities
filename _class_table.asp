<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Class Table
	Public Attr
	Public Rows()
	
	Public Sub Class_Initialize()
		ReDim Rows(0)
	End Sub
	
	Public Function NewRow(ByVal att)
		Set NewRow = New Row
		NewRow.Attr = att
		' since top element is never used, we can place and object in it
		' before we increase the size of the array
		Set Rows(UBound(Rows)) = NewRow
		ReDim Preserve Rows(UBound(Rows)+1)
	End Function
	
	Public Sub Output()
		Dim i
		Response.Write "<table " & Attr & ">"
		For i = 0 To UBound(Rows)-1
			Rows(i).Output()
		Next
		Response.Write "</table>"
	End Sub
End Class

Function NewTable ( ByVal att )
	Set NewTable = New Table
	NewTable.Attr = att
End Function

Class Row
	Public Attr
	Public Cells()
	
	Public Sub Class_Initialize()
		ReDim Cells(0)
	End Sub

	Public Function NewHeader(ByVal att)
		Set NewHeader = NewCell(att)
		NewHeader.Tag = "th"
	End Function
	
	Public Function NewCell(ByVal att)
		Set NewCell = New Cell
		NewCell.Attr = att
		' since top element is never used, we can place and object in it
		' before we increase the size of the array
		Set Cells(UBound(Cells)) = NewCell
		ReDim Preserve Cells(UBound(Cells)+1)
	End Function
	
	Public Sub Output()
		Dim i
		Response.Write "<tr " & Attr & ">"
		For i = 0 To UBound(Cells)-1
			Cells(i).Output()
		Next
		Response.Write "</tr>"
	End Sub
End Class

Class Cell
	Public Tag
	Public Attr
	Public Value

	Public Sub Class_Initialize()
		Tag = "td"
	End Sub
	
	Public Sub Output()
		Response.Write "<" & Tag & " " & Attr & ">" & Value & "</" & Tag & ">"
	End Sub
End Class
%>