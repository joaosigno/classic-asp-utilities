<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function Rot47 ( ByVal s )
	Dim i, last, b, n
	b = StrConv ( s, vbFromUnicode )
	last = UBound(b)
	For i = LBound(b) To last
		n = b(i)
		If n >= 33 And n <= 126 Then
			b(i) = (n + 14) Mod 94 + 33
		End If
	Next
	Rot47 = StrConv ( b, vbUnicode )
End Function
Sub Rot47_Test ( ByVal orig, ByVal expect )
	Dim x
	x = Rot47(orig)
	If x <> expect Then
		Err.Raise -1, "Rot47", "Encoding """ & orig & """, expected """ & expect & """, but got """ & x & """"
	End If
End Sub
Sub Rot47_Tests()
	Rot47_Test "Hello, world!", "w6==@[ H@C=5P"
End Sub
Call Rot47_Tests()
%>
