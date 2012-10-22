<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function MDBOpenEx(ByVal FileName, ByVal UserName, ByVal Password)
	On Error GoTo 0
	Set MDBOpenEx = CreateObject("ADODB.Connection")
	MDBOpenEx.ConnectionTimeout = 20
	On Error Resume Next
	Err.Clear
	MDBOpenEx.Open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & FileName & ";Uid=" & UserName & ";Pwd=" & Password & ";"
	If Err Then
		Dim sErr
		sErr = Err.Number & " - " & Err.Description
		On Error GoTo 0
		Err.Raise -1, "MDBOpenEx()", "Error occurred opening database """ & FileName & """: " & sErr
		Set MDBOpenEx = Nothing
	End If
End Function

Function MDBOpen(ByVal FileName)
	Set MDBOpen = MDBOpenEx(FileName,"Admin","")
End Function
%>
