<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function MSSQLOpen ( ByRef con, ByVal server, ByVal db, ByVal user, ByVal pass )
	MSSQLOpen = False ' default to failure unless we make it to the end of the function
	On Error Resume Next
	Err.Clear
	con.Open "Provider=SQLNCLI.1;" & _
		"Data Source=" & server & ";" & _
		"User ID=" & user & ";" & _
		"Password=" & pass & ";" & _
		"Initial Catalog=" & db & ";" & _
		"Persist Security Info=False;"
	'con.Open "Provider=sqloledb;" & _ 
	'	"Data Source=" & server & ";" & _
	'	"Initial Catalog=" & db & ";" & _
	'	"Network=DBMSSOCN;" & _
	'	"User Id=" & user & ";" & _
	'	"Password=" & pass
	If 0=Err.Number Then MSSQLOpen = True
	On Error GoTo 0
End Function
%>
