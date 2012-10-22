<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Class FileIO
	Private fso, ts

	Private Sub Class_Initialize()
		Set fso = CreateObject("Scripting.FileSystemObject")
	End Sub

	' open an existing file for input
	Public Function OpenInput ( ByVal FileName )
		OpenInput = False
		If Not fso.FileExists(FileName) Then Exit Function
		Set ts = fso.OpenTextFile ( FileName, ForReading )
		OpenInput = True
	End Function

	' open a new file for writing, chokes on existing file
	Public Function OpenOutput ( ByVal FileName, ByVal bOverwrite )
		OpenOutput = False
		If Not bOverwrite And fso.FileExists(FileName) Then Exit Function
		On Error Resume Next
		Err.Clear
		Set ts = fso.OpenTextFile ( FileName, ForWriting, True, TristateFalse )
		If 0 = Err.Number Then OpenOutput = True
	End Function

	' open a file for appending
	' if it doesn't exist and bCreate is true, creates the file
	Public Function OpenAppend ( ByVal FileName, ByVal bCreate )
		OpenAppend = False
		If Not bCreate And Not fso.FileExists(FileName) Then Exit Function
		Set ts = fso.OpenTextFile ( FileName, ForAppending, bCreate, TristateFalse )
		OpenAppend = True
	End Function

	Public Function Close()
		ts.Close
	End Function

	' ###########################################################################
	' input functions...
	' ###########################################################################

	Public Property Get Eof()
		Eof = ts.AtEndOfStream
	End Property

	Public Function NextLine ( ByRef line )
		NextLine = False
		If Eof Then Exit Function
		line = ts.ReadLine()
		If line <> "" Or Not Eof Then NextLine = True
	End Function

	' ###########################################################################
	' output functions
	' ###########################################################################
	Public Sub Print ( ByVal text )
		ts.Write text
	End Sub
	Public Sub Write ( ByVal text )
		ts.Write text
	End Sub
	Public Sub PrintLine ( ByVal text )
		ts.WriteLine text
	End Sub
	Public Sub WriteLine ( ByVal text )
		ts.WriteLine text
	End Sub

End Class
%>
