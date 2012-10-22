<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function CSafeLng ( ByVal x )
	CSafeLng = CSafeLngEx ( x, 0 )
End Function
Function CSafeLngEx ( ByVal x, ByVal def )
	On Error Resume Next
	Err.Clear
	CSafeLngEx = CLng(x)
	If Err Then
		Err.Clear
		CSafeLngEx = def
	End If
	On Error GoTo 0
End Function
%>
