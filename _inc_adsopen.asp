<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Public Function ADSOpen(ByVal DirName)
	Set ADSOpen = ADSOpen2 ( DirName, "", "" )
End Function

Public Function ADSOpenEx(ByVal DirName, ByVal bServer)
	Set ADSOpenEx = ADSOpen2Ex ( DirName, "", "", bServer )
End Function

Public Function ADSOpen2(ByVal DirName, ByVal Username, ByVal Password)
	'Response.Write "ADSOpen(" & DirName & ")<br>"
	'On Error Resume Next
	Err.Clear
	Set ADSOpen2 = ADSOpen2Ex(DirName, Username, Password, True)
	If Err.Number = 80004005 Then
		On Error GoTo 0
		Set ADSOpen2 = ADSOpen2Ex(DirName, Username, Password, False)
	ElseIf Err.Number<>0 Then
		Dim nErr, sErr
		nErr = CLng(Err.Number)
		sErr = CStr(Err.Description)
		On Error GoTo 0
		Err.Raise nErr, "ADSOpen2()", sErr
	End If
End Function

Public Function ADSOpen2Ex(ByVal DirName, ByVal Username, ByVal Password, ByVal bServer)
	Dim stype
	'If bServer Then
		stype = "ADS_REMOTE_SERVER"
	'Else
	'	stype = "ADS_LOCAL_SERVER"
	'End If
	Set ADSOpen2Ex = CreateObject("ADODB.Connection")

	'ADSOpenEx.Open "Provider=Advantage.OLEDB.1;Data Source=" & DirName & ";ServerType=" & stype & ";LockMode=ADS_COMPATIBLE_LOCKING;TableType=ADS_CDX"
	ADSOpen2Ex.Open "Provider=Advantage.OLEDB.1;Password=""" & Password & """;User ID=""" & Username & """;Data Source=" & DirName & ";ServerType=" & stype & ";LockMode=ADS_COMPATIBLE_LOCKING;TableType=ADS_CDX"
End Function

Public Function ADTOpen(DirName)
	'On Error Resume Next
	'Err.Clear
	Set ADTOpen = ADTOpenEx(DirName, True)
	'If Err.Number = 80004005 Then
	'	On Error GoTo 0
	'	Set ADTOpen = ADTOpenEx(DirName, False)
	'End If
End Function

Public Function ADTOpenEx(DirName, bServer)
	Dim stype
	'If bServer Then
		stype = "ADS_REMOTE_SERVER"
	'Else
	'	stype = "ADS_LOCAL_SERVER"
	'End If
	Set ADTOpenEx = CreateObject("ADODB.Connection")

	ADTOpenEx.Open "Provider=Advantage.OLEDB.1;Data Source=" & DirName & ";ServerType=" & stype & ";LockMode=ADS_COMPATIBLE_LOCKING;TableType=ADS_ADT"
End Function

Public Function ADTDate ( ByVal d )
	ADTDate = Year(d) _
		& "-" & Right("0"&Month(d),2) _
		& "-" & Right("0"&Day(d),2) _
		& " " & Right("0"&Hour(d),2) _
		& ":" & Right("0"&Minute(d),2) _
		& ":" & Right("0"&Second(d),2)
End Function
%>