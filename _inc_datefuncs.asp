<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function DateOnly ( ByVal d )
	If IsDate(d) Then
		DateOnly = DateSerial ( Year(d), Month(d), Day(d) )
	Else
		DateOnly = ""
	End If
End Function

Function TimeOnly ( ByVal d )
	TimeOnly = TimeSerial ( Hour(d), Minute(d), Second(d) )
End Function

Function YYYYMMDD ( ByVal d, ByVal sep )
	YYYYMMDD = Year(d) & sep & Right("0" & Month(d),2) & sep & Right("0" & Day(d),2)
End Function

Function LastDOM ( ByVal d )
	d = DateSerial(Year(d),Month(d),28)+4
	LastDOM = DateSerial(Year(d),Month(d),1)-1
End Function
%>