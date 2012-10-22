<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function String2Binary(sString)
	Dim i
	For i = 1 to Len(sString)
	   String2Binary = String2Binary & ChrB(AscB(Mid(sString,i,1)))
	Next
End Function

'Byte string to string conversion
Function Binary2String(bsString)
	Dim i
	Binary2String = ""
	For i = 1 to LenB(bsString)
	   Binary2String = Binary2String & Chr(AscB(MidB(bsString,i,1))) 
	Next
End Function
%>
