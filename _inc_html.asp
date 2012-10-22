<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/

Function ToHtml ( ByVal s )
	ToHtml = Server.HtmlEncode(s)
End Function
Function ToUrl ( ByVal s )
	ToUrl = Server.URLEncode(s)
End Function

%>