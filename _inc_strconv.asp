<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Const vbFromUnicode = 128
Const vbUnicode = 64

Public Function StrConv ( stringData, ByVal conversion )
	Dim tmp, i, last
	Select Case conversion
	Case vbFromUnicode
		' Converts a Unicode string to Ascii byte array
		last = Len(stringData)
		ReDim tmp(last-1)
		For i = 1 To last
			tmp(i-1) = Asc(Mid(stringData,i,1))
		Next
		StrConv = tmp
	Case vbUnicode
		' Converts an Ascii byte array to string
		last = UBound(stringData)
		tmp = ""
		StrConv = ""
		For i = LBound(stringData) To last
			tmp = tmp & Chr(stringData(i))
			If Len(tmp) >= 256 Then
				StrConv = StrConv & tmp
				tmp = ""
			End If 
		Next
		If Len(tmp) > 0 Then
			StrConv = StrConv & tmp
		End If
	End Select
End Function
%>