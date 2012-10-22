<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function Cut ( ByVal text, ByVal sep, ByVal col )
	Dim ar, i
	ar = Split ( text, sep )
	If col > 0 Then
		i = col - 1
	ElseIf col < 0 Then
		i = UBound(ar) + col + 1
	Else
		i = -1
	End If
	If i < 0 Or i > UBound(ar) Then
		Cut = ""
	Else
		Cut = ar(i)
	End If
End Function
%>
