<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function Rot13 ( ByVal s )
	Dim i, last, c, n
	last = Len(s)
	For i = 1 To last
		c = Mid(s,i,1)
		Select Case c
		Case 65 To 90 ' upper-case
			Mid(s,i,1) = (Asc(c) - 52) Mod 26 + 65
		Case 97 To 122
			Mid(s,i,1) = (Asc(c) - 84) Mod 26 + 97
		End Select
	Next i
	Rot13 = s
End Function
Function Rot47 ( ByVal s )
	Dim i, last, c, n
	last = Len(s)
	For i = 1 To last
		c = Mid(s,i,1)
		Select Case c
		Case 33 To 126
			Mid(s,i,1) = (Asc(c) + 14) Mod 94 + 33
		End Select
	Next i
	Rot13 = s
End Function
Sub Rot13_Tests()
	Rot13_Test "@AZ[`az{", "@NM[`nm{"
End Sub
Call Rot13_Tests()
%>
