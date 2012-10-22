<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function SafeDiv(n,d)
	On Error Resume Next
	Err.Clear
	SafeDiv = n/d
	If Err Then SafeDiv = 0
	On Error GoTo 0
End Function

Function Min ( ByVal a, ByVal b )
	If a < b Then
		Min = a
	Else
		Min = b
	End If
End Function

Function Max ( ByVal a, ByVal b )
	If a > b Then
		Max = a
	Else
		Max = b
	End If
End Function
%>