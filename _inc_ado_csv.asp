<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function OpenAdoCsv ( ByVal Path )
	Set OpenAdoCsv = CreateObject("ADODB.Connection")
	OpenAdoCsv.Open "Driver={Microsoft Text Driver (*.txt; *.csv)};" & _
		"Dbq=" & Path & ";Extensions=asc,csv,tab,txt;Persist Security Info=False"
End Function
%>