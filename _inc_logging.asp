<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function OpenLog ( ByVal Path )
	Dim fso
	Set fso = CreateObject("Scripting.FileSystemObject")
	If fso.FileExists(Path) Then
		Set OpenLog = fso.OpenTextFile ( Path, ForAppending )
	Else
		Set OpenLog = fso.CreateTextFile ( Path, False )
	End If
End Function

Sub AppendLog ( ByVal Path, ByVal msg )
	OpenLog(Path).WriteLine msg
End Sub
%>