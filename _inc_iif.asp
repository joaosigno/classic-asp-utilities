<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function IIf ( ByVal cond, ByRef vTrue, ByRef vFalse )
	If cond Then
		If IsObject(vTrue) Then
			Set IIf = vTrue
		Else
			IIf = vTrue
		End If
	Else
		If IsObject(vFalse) Then
			Set IIf = vFalse
		Else
			IIf = vFalse
		End If
	End If
End Function
%>