<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Public Function MysqlOpen(ByRef con, ByVal sServer, ByVal sDatabase, ByVal sUserName, ByVal sPassword)
    On Error Resume Next
    Err.Clear
    con.ConnectionTimeout = 20
    'con.CursorLocation = adUseClient ' causes MoveNext to have troubles...
    con.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
        & "SERVER=" & sServer & ";" _
        & "DATABASE=" & sDatabase & ";" _
        & "UID=" & sUserName & ";" _
        & "PWD=" & sPassword & ";" _
        & "OPTION=" & (1 + 2 + 8 + 32 + 2048 + 16384) 'SET ALL PARAMETERS ( 2 = return matching rows, 8 = allow BIG results, 16384 = no BIGINT support )
    con.Open
    MysqlOpen = (0 = Err.Number)
End Function
%>