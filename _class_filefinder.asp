<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Class FileFinder
	Public Sub Class_Initialize()
	End Sub
	
	Public Function Find ( ByVal FileSpec )
		Dim d, sh, x, line
		Set d = New Dictionary
		Set sh = CreateObject("WScript.Shell")
		Set x = sh.Exec("cmd /C dir /b " & FileSpec)
		'Response.Write "FileFinder.Find(" & FileSpec & ")<br>"

		While Not x.StdOut.AtEndOfStream
			line = x.StdOut.ReadLine()
			'Response.Write line & "<br>"
			'fileNameMatch = fileNameRegex.exec(directoryContentLine);
			d(line) = line
		Wend
		'Response.Write "d.Count=" & d.Count & "<br>"
		Set Find = d
	End Function
End Class
%>
