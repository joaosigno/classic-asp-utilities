<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function RoundFixed ( Byval f, ByVal digits )
	f = CStr ( Round ( f, digits ) )
	If digits > 0 Then
		If 0 = InStr ( f, "." ) Then f = f & "."
		f = f & String(digits,"0")
		f = Left ( f, InStr(f,".") + digits )
	End If
	RoundFixed = f
End Function
%>
