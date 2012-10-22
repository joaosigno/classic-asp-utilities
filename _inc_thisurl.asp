<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Public Function ThisUrl()
	Dim tmp
	ThisUrl = Trim(Request.ServerVariables("URL"))
	tmp = InStrRev(ThisUrl,"/")
	If tmp > 0 Then ThisUrl = Mid(ThisUrl,tmp+1)
	tmp = Trim(Request.ServerVariables("QUERY_STRING"))
	If tmp <> "" Then ThisUrl = ThisUrl & "?" & tmp
End Function
%>
