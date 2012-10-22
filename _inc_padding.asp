<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/

Function RPad ( ByVal s, ByVal n )
	Dim lenS
	If IsNull(s) Then
		lenS = 0
	Else
		lenS = Len(s)
	End If
	If lenS < n Then
		RPad = s & Space(n-lenS)
	Else
		RPad = s
	End If
End Function

Function LPad ( ByVal s, ByVal n )
	Dim lenS
	If IsNull(s) Then
		lenS = 0
	Else
		lenS = Len(s)
	End If
	If lenS < n Then
		LPad = Space(n-lenS) & s
	Else
		LPad = s
	End If
End Function

Function RPadClamp ( ByVal s, ByVal n )
	s = Left(s,n)
	Dim lenS
	If IsNull(s) Then
		lenS = 0
	Else
		lenS = Len(s)
	End If
	If lenS < n Then
		RPadClamp = s & Space(n-lenS)
	Else
		RPadClamp = s
	End If
End Function

Function LPadClamp ( ByVal s, ByVal n )
	s = Left(s,n)
	Dim lenS
	If IsNull(s) Then
		lenS = 0
	Else
		lenS = Len(s)
	End If
	If lenS < n Then
		LPadClamp = Space(n-lenS) & s
	Else
		LPadClamp = s
	End If
End Function
%>
