<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Const SQL_TRUE = "1"
Const SQL_FALSE = "0"

Const dbfOpenUnspecified = -1 ' adOpenUnspecified
Const dbfOpenForwardOnly = 0 ' adOpenForwardOnly
Const dbfOpenKeyset = 1 ' adOpenKeyset
Const dbfOpenDynamic = 2 ' adOpenDynamic
Const dbfOpenStatic = 3 ' adOpenStatic

Const dbfLockReadOnly = 1 ' adLockReadOnly
Const dbfLockPessimistic = 2 ' adLockPessimistic
Const dbfLockOptimistic = 3 ' adLockOptimistic

Function NewCon()
	Set NewCon = CreateObject("ADODB.Connection")
End Function

Function NewRS ( ByRef con ) ' con is optional, pass Null to avoid it
	Set NewRS = CreateObject("ADODB.Recordset")
	NewRS.CursorLocation = 3 ' adUseClient
	If IsObject(con) And Not IsNull(con) Then
		Set NewRS.ActiveConnection = con
	End If
End Function
Function InitRS ( ByRef rs, ByVal LockType )
	rs.CursorType = dbfOpenStatic ' used to be dbfOpenDynamic
	rs.LockType = LockType
	Set InitRS = rs
End Function
Function InitRSOptimistic ( ByRef rs )
	Set InitRSOptimistic = InitRS(rs,dbfLockOptimistic)
End Function
Function InitRSPessimistic ( ByRef rs )
	Set InitRSPessimistic = InitRS(rs,dbfLockPessimistic)
End Function
Function InitRSReadOnly ( ByRef rs )
	Set InitRSReadOnly = InitRS(rs,dbfLockReadOnly)
End Function
Function NewRSOptimistic ( ByRef con ) ' con is optional, pass Null to avoid it
	Set NewRSOptimistic = InitRSOptimistic(NewRS(con))
End Function
Function NewRSPessimistic ( ByRef con ) ' con is optional, pass Null to avoid it
	Set NewRSPessimistic = InitRSPessimistic(NewRS(con))
End Function
Function NewRSReadOnly ( ByRef con ) ' con is optional, pass Null to avoid it
	Set NewRSReadOnly = InitRSReadOnly(NewRS(con))
End Function
Function NewRSInit ( ByRef con, ByVal LockType )
	Set NewRSInit = InitRS ( NewRS(con), LockType )
End Function

' alternate naming convention for the locking:

Const dbfLockShared = 3 ' adLockOptimistic
Const dbfLockExclusive = 2 ' adLockPessimistic

Function InitRSShared ( ByRef rs )
	Set InitRSShared = InitRS(rs,dbfLockShared)
End Function
Function InitRSExclusive ( ByRef rs )
	Set InitRSShared = InitRS(rs,dbfLockExclusive)
End Function
Function NewRSShared ( ByRef con )
	Set NewRSShared = InitRSShared(NewRS(con))
End Function
Function NewRSExclusive ( ByRef con )
	Set NewRSExclusive = InitRSExclusive(NewRS(con))
End Function

Function ConExec ( ByRef con, ByVal sql, ByRef sErr )
	sErr = ""
	On Error Resume Next
	Err.Clear
	Set ConExec = con.Execute ( sql )
	If Err.Number <> 0 Then
		sErr = Err.Description
		Set ConExec = Nothing
	End If
	On Error GoTo 0
End Function

Function RsOpen ( ByRef rs, ByVal sql, ByRef sErr )
	On Error Resume Next
	Err.Clear
	rs.Open sql
	If Err.Number <> 0 Then
		sErr = Err.Description
		RsOpen = False
	Else
		sErr = ""
		RsOpen = True
	End If
	On Error GoTo 0
End Function

Function RsOpenEof ( ByRef rs, ByVal sql, ByRef sErr )
	RsOpenEof = RsOpen ( rs, sql, sErr )
	If RsOpenEof Then
		If rs.EOF Then
			sErr = "EOF"
			RsOpenEof = False
			rs.Close
		End If
	End If
End Function

' UpdateAndRequery is used to obtain the value of an autonumber field
Sub UpdateAndRequery ( ByRef rs )
	rs.Update
	Dim bm : bm = rs.AbsolutePosition
	rs.Requery
	rs.AbsolutePosition = bm
End Sub

Function SimpleQuery ( con, sql )
	Dim ar(), rs
	ReDim ar(0)
	Set rs = NewRSReadOnly(con)
	On Error Resume Next
	Err.Clear
	rs.Open sql
	If Err.Number <> 0 Then
		Dim nErr, sErr
		nErr = Err.Number
		sErr = Err.Description
		On Error GoTo 0
		DBFuncs_Err "ModDBFuncs.SimpleQuery()", "Error " & nErr & ": " & sErr & vbCrLf & "SQL:" & vbCrLf & sql
	End If
	On Error GoTo 0
	While Not rs.EOF
		ar(UBound(ar)) = rs.Fields(0).Value
		ReDim Preserve ar(UBound(ar)+1)
		rs.MoveNext
	Wend
	rs.Close
	Set rs = Nothing
	If UBound(ar) = 1 Then
		SimpleQuery = ar(0)
	Else
		ReDim Preserve ar(UBound(ar)-1)
		SimpleQuery = ar
	End If
End Function

Function RsFieldList ( ByRef rs )
	Dim i, ar()
	ReDim ar(rs.Fields.Count-1)
	For i = 0 To rs.Fields.Count-1
		ar(i) = rs(i).Name
	Next
	RsFieldList = ar
End Function

Function TableFieldList ( ByRef con, ByVal TableName )
	Dim rs : Set rs = NewRSReadOnly(con)
	rs.Open TableName
	TableFieldList = RsFieldList(rs)
End Function

Function InQuote ( ByVal str, ByVal idx, ByVal q, ByVal quoted_q )
	Dim i
	InQuote = False
	For i = 1 To idx-1
		If Mid(str,i,Len(quoted_q)) = quoted_q Then
			i = i + Len(quoted_q) - 1
		ElseIf Mid(str,i,Len(q)) = q Then
			InQuote = Not InQuote
			i = i + Len(q) - 1
		End If
	Next
End Function

Function SqlParameterInject ( ByVal sql, ByVal parameters )
	If Not IsArray(parameters) Then parameters = Array(parameters)
	Dim s, i, j, tmp
	j = 0
	s = sql
	For i = 0 To UBound(parameters)
		j = InStr(j+1,s,"?")
		If j = 0 Then Exit For
		Select Case TypeName(parameters(i))
		Case "Empty"
			tmp = "''"
		Case "String"
			tmp = "'" & Replace(parameters(i),"'","''") & "'"
		Case Else
			DBFuncs_Err "ModDBFuncs.SqlParameterInject()", "parameters #{0} unknown type '{1}', sql=<<{2}>>, parameters=<<{3}>>", Array ( i, TypeName(parameters(i)), sql, Join(parameters,",") )
			SqlParameterInject = Empty : Exit Function
		End Select
		s = Left(s,j-1) & tmp & Mid(s,j+1)
		j = j + Len(tmp) - 1 ' don't find ? in replaced text
	Next
	DBFuncs_Trace "ModDBFuncs.SqlParameterInject()", "final sql: {0}", s
	SqlParameterInject = s
End Function

Function SqlOpenFEx ( ByRef rs, ByVal sql, ByVal parameters )
	DBFuncs_Trace "ModDBFuncs.SqlOpenFEx()", "sql=<<{0}>>, parameters={1}", Array ( sql, Join(parameters,",") )
	sql = SqlParameterInject ( sql, parameters )
	rs.Open sql
	SqlOpenFEx = True
End Function

Function SqlExecFEx ( ByRef con, ByVal sql, ByVal parameters )
	Dim rs
	Set rs = NewRSReadOnly(con)
	If SqlOpenFEx ( rs, sql, parameters ) Then
		Set SqlExecFEx = rs
	Else
		Set SqlExecFEx = Nothing
	End If
End Function

Function ADOOpenFEx ( ByRef rs, ByVal sql, ByVal parameters )
	DBFuncs_Trace "ModDBFuncs.ADOOpenFEx()", "sql=<<{0}>>, parameters=<<{1}>>", Array ( sql, Join(parameters,",") )
	If Not IsArray(parameters) Then parameters = Array(parameters)
	Dim cmd, i, nErr, sErr
	Set cmd = CreateObject("ADODB.Command")
	cmd.CommandType = adCmdText
	cmd.CommandText = sql
	cmd.ActiveConnection = rs.ActiveConnection
	For i = 0 To UBound(parameters)
		Dim param, dte, n
		Select Case TypeName(parameters(i))
			Case "Empty"
				DBFuncs_Trace "ModDBFuncs.ADOOpenFEx()", "Parameter #{0} is empty", i ' Don't raise an error here, makes it harder to track down. If we return nothing, our calling script will surely bomb :)
				Set ADOOpenFEx = Nothing
				Exit Function
			Case "String"
				DBFuncs_Trace "ModDBFuncs.ADOOpenFEx()", "Parameter #{0} is string", i
				dte = adVarChar
				n = Len(parameters(i))
			Case Else
				DBFuncs_Err "ModDBFuncs.ADOOpenFEx()", "Unsupported type: '{0}' for sql=<<{1}>>, params=<<{2}>>", Array ( TypeName(parameters(i)), sql, Join(parameters,",") )
				ADOOpenFEx = False
				Exit Function
		End Select
		Err.Clear
		Set param = cmd.CreateParameter("@" & i, _
								dte, _
								adParamInput, _
								n, _
								parameters(i))
		'With param
		'	DBClass_Trace "param.Name=" & .Name
		'	DBClass_Trace "param.Type=" & .Type
		'	DBClass_Trace "param.Direction=" & .Direction & "(adParamInput=" & adParamInput & ")"
		'	DBClass_Trace "param.Value=" & .Value
		'End With
		On Error Resume Next
		Err.Clear
		cmd.Parameters.Append param
		If Err Then
			nErr = Err.Number
			sErr = Err.Description
			On Error GoTo 0
			DBFuncs_Err "ModDBFuncs.ADOOpenFEx()", "Error occurred trying to add parameter #{0}: ({1}) {2}", Array ( i, nErr, sErr )
		End If
		On Error GoTo 0
	Next
	On Error Resume Next
	Err.Clear
	rs.Open cmd
	If Err Then
		sErr = Err.Number & " - " & Err.Description
		On Error GoTo 0
		Err.Raise -1, "ADOOpenFEx()", "Error: " & sErr & vbCrLf & "cmd: " & cmd.CommandText
	End If
	ADOOpenFEx = True
End Function

Function ADOExecFEx(ByRef con, ByVal sql, ByVal parameters)
	Dim rs
	Set rs = NewRSReadOnly(con)
	If Not ADOOpenFEx(rs, sql, parameters) Then Set rs = Nothing
	Set ADOExecFEx = rs
End Function

Function ADOExecFSafe ( ByRef con, ByVal sql, ByRef parameters )
	Dim sql2, nErr, sErr
	On Error Resume Next
	Err.Clear
	sql2 = SqlFormat ( sql, parameters )
	If Err Then
		nErr = Err.Number
		sErr = Err.Description
		On Error GoTo 0
		DBFuncs_Err "ModDBFuncs.ADOExecFSafe()", "Error {0}: {1}{br}trying to format SQL:{br}{2}", Array ( nErr, sErr, sql )
	End If
	Err.Clear
	Set ADOExecFSafe = ADOExecFEx ( con, sql2, Null )
	If Err Then
		nErr = Err.Number
		sErr = Err.Description
		On Error GoTo 0
		DBFuncs_Err "ModDBFuncs.ADOExecFSafe()", "Error {0}: {1}{br}SQL:{br}{2}", Array(nErr,sErr,sql2)
	End If
End Function

Function KeyBefore ( con, table, keyName, dte, keyVal )
	Dim rs
	Set rs = NewRSReadOnly()
	rs.Open "select top 1 [" & keyName & "] from [" & table & "] where [" & keyName & "] < " & SQLWrap(dte,keyVal) & " order by [" & keyName & "] DESC", con
	If Not rs.EOF Then
		KeyBefore = rs(keyName).Value
	Else
		KeyBefore = ""
	End If
	rs.Close
	Set rs = Nothing
End Function

Function KeyAfter ( con, table, keyName, dte, keyVal )
	Dim rs
	Set rs = NewRSReadOnly()
	rs.Open "select top 1 [" & keyName & "] from [" & table & "] where [" & keyName & "] > " & SQLWrap(dte,keyVal) & " order by [" & keyName & "]", con
	If Not rs.EOF Then
		KeyAfter = rs(keyName).Value
	Else
		KeyAfter = ""
	End If
	rs.Close
	Set rs = Nothing
End Function

Private Function SQLIsDate(field)
	Select Case field

	Case adDate, adDBDate, adDBTime, adDBTimeStamp
		SQLIsDate = True

	Case Else
		SQLIsDate = False

	End Select
End Function

Private Function SQLFieldSize(field)
	Select Case field.Type
		Case adDate, adDBDate, adDBTime, adDBTimeStamp

			SQLFieldSize = Len("12/31/9999 11:59:59 PM")

		Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt,adUnsignedSmallInt,adUnsignedTinyInt,adVarNumeric

			SQLFieldSize = field.Precision ' DefinedSize * 3 ???

		Case Else

			SQLFieldSize = field.DefinedSize
	End Select
End Function

Private Function DataTypeDesc(dte)
	Select Case dte
	Case adArray
		DataTypeDesc = "adArray"
	Case adBigInt
		DataTypeDesc = "adBigInt"
	Case adBinary
		DataTypeDesc = "adBinary"
	Case adBoolean
		DataTypeDesc = "adBoolean"
	Case adBSTR
		DataTypeDesc = "adBSTR"
	Case adChapter
		DataTypeDesc = "adChapter"
	Case adChar
		DataTypeDesc = "adChar"
	Case adCurrency
		DataTypeDesc = "adCurrency"
	Case adDate
		DataTypeDesc = "adDate"
	Case adDBDate
		DataTypeDesc = "adDBDate"
	Case adDBTime
		DataTypeDesc = "adDBTime"
	Case adDBTimeStamp
		DataTypeDesc = "adDBTimeStamp"
	Case adDecimal
		DataTypeDesc = "adDecimal"
	Case adDouble
		DataTypeDesc = "adDouble"
	Case adEmpty
		DataTypeDesc = "adEmpty"
	Case adError
		DataTypeDesc = "adError"
	Case adFileTime
		DataTypeDesc = "adFileTime"
	Case adGUID
		DataTypeDesc = "adGUID"
	Case adIDispatch
		DataTypeDesc = "adIDispatch"
	Case adInteger
		DataTypeDesc = "adInteger"
	Case adIUnknown
		DataTypeDesc = "adIUnknown"
	Case adLongVarBinary
		DataTypeDesc = "adLongVarBinary"
	Case adLongVarChar
		DataTypeDesc = "adLongVarChar"
	Case adLongVarWChar
		DataTypeDesc = "adLongVarWChar"
	Case adNumeric
		DataTypeDesc = "adNumeric"
	Case adPropVariant
		DataTypeDesc = "adPropVariant"
	Case adSingle
		DataTypeDesc = "adSingle"
	Case adSmallInt
		DataTypeDesc = "adSmallInt"
	Case adTinyInt
		DataTypeDesc = "adTinyInt"
	Case adUnsignedBigInt
		DataTypeDesc = "adUnsignedBigInt"
	Case adUnsignedInt
		DataTypeDesc = "adUnsignedInt"
	Case adUnsignedSmallInt
		DataTypeDesc = "adUnsignedSmallInt"
	Case adUnsignedTinyInt
		DataTypeDesc = "adUnsignedTinyInt"
	Case adUserDefined
		DataTypeDesc = "adUserDefined"
	Case adVarBinary
		DataTypeDesc = "adVarBinary"
	Case adVarChar
		DataTypeDesc = "adVarChar"
	Case adVariant
		DataTypeDesc = "adVariant"
	Case adVarNumeric
		DataTypeDesc = "adVarNumeric"
	Case adVarWChar
		DataTypeDesc = "adVarWChar"
	Case adWChar
		DataTypeDesc = "adWChar"
	Case Else
		DataTypeDesc = CLng(t) & " (?Unknown)"
    End Select
End Function

Function FixDbValue ( ByVal value )
	' This point of this function is to detect value types that aren't supported
	' by vbscript and convert them to a compatible useable type

	Select Case VarType(value)
	Case 14 ' decimal
		FixDbValue = CCur(value)
	Case Else
		FixDbValue = value
	End Select

End Function

Function SQLIsChar ( dte )
	SQLIsChar = InStr ( LCase(DataTypeDesc(dte)), "char" ) > 0
End Function
Function SQLQuote ( v )
	SQLQuote = SQLWrap ( adVarChar, v )
End Function
Function SQLWrap ( dte, v )
	If Len(v) = 0 Or IsNull(v) Or IsEmpty(v) Then
		SQLWrap = "Null"
		Exit Function
	End If
	Select Case dte
	'Case adArray
	'    SQLWrap = ?
	Case adBigInt
		SQLWrap = v
	'Case adBinary
	'    SQLWrap = ?
	Case adBoolean
		SQLWrap = v
	Case adBSTR
		SQLWrap = "'" & Replace(v,"'","''") & "'"
	'Case adChapter
	'    SQLWrap = ?
	Case adChar
		SQLWrap = "'" & Replace(v,"'","''") & "'"
	Case adCurrency
		SQLWrap = v
	Case adDate
		SQLWrap = "#" & v & "#"
	Case adDBDate
		SQLWrap = "#" & v & "#"
	Case adDBTime
		SQLWrap = "#" & v & "#"
	Case adDBTimeStamp
		SQLWrap = "#" & v & "#"
	Case adDecimal
		SQLWrap = v
	Case adDouble
		SQLWrap = v
	'Case adEmpty
	'    SQLWrap = ?
	'Case adError
	'    SQLWrap = ?
	Case adFileTime
		SQLWrap = "#" & v & "#"
	'Case adGUID
	'    SQLWrap = ?
	'Case adIDispatch
	'    SQLWrap = ?
	Case adInteger
		SQLWrap = v
	'Case adIUnknown
	'    SQLWrap = ?
	'Case adLongVarBinary
	'    SQLWrap = ?
	Case adLongVarChar
		SQLWrap = "'" & Replace(v,"'","''") & "'"
	Case adLongVarWChar
		SQLWrap = "'" & Replace(v,"'","''") & "'"
	Case adNumeric
		SQLWrap = v
	'Case adPropVariant
	'    SQLWrap = ?
	Case adSingle
		SQLWrap = v
	Case adSmallInt
		SQLWrap = v
	Case adTinyInt
		SQLWrap = v
	Case adUnsignedBigInt
		SQLWrap = v
	Case adUnsignedInt
		SQLWrap = v
	Case adUnsignedSmallInt
		SQLWrap = v
	Case adUnsignedTinyInt
		SQLWrap = v
	'Case adUserDefined
	'    SQLWrap = ?
	'Case adVarBinary
	'    SQLWrap = ?
	Case adVarChar
		SQLWrap = "'" & Replace(v,"'","''") & "'"
	'Case adVariant
	'    SQLWrap = ?
	Case adVarNumeric
		SQLWrap = v
	Case adVarWChar
		SQLWrap = "'" & Replace(v,"'","''") & "'"
	Case adWChar
		SQLWrap = "'" & Replace(v,"'","''") & "'"
	Case Else
		DBFuncs_Err "ModDBFuncs.SQLWrap()", "Don't know how to wrap ADODB.DataTypeEnum: {0}", DataTypeDesc(dte)
	End Select
End Function
Function SQLUnwrap ( ByRef fld )
	Select Case fld.Type
	'Case adArray
	'    SQLUnwrap = ?
	Case adBigInt
		SQLUnwrap = fld.Value
	'Case adBinary
	'    SQLUnwrap = ?
	Case adBoolean
		SQLUnwrap = fld.Value
	'Case adBSTR
	'    SQLUnwrap = ? "'" & Replace(v,"'","''") & "'"
	'Case adChapter
	'    SQLUnwrap = ?
	Case adChar
		SQLUnwrap = fld.Value
	Case adCurrency
		SQLUnwrap = fld.Value
	Case adDate
		SQLUnwrap = fld.Value
	Case adDBDate
		SQLUnwrap = fld.Value
	Case adDBTime
		SQLUnwrap = fld.Value
	Case adDBTimeStamp
		SQLUnwrap = fld.Value
	Case adDecimal
		SQLUnwrap = fld.Value
	Case adDouble
		SQLUnwrap = fld.Value
	Case adEmpty
		SQLUnwrap = fld.Value
	'Case adError
	'    SQLUnwrap = ?
	Case adFileTime
		SQLUnwrap = fld.Value
	'Case adGUID
	'    SQLUnwrap = ?
	'Case adIDispatch
	'    SQLUnwrap = ?
	Case adInteger
		SQLUnwrap = fld.Value
	'Case adIUnknown
	'    SQLUnwrap = ?
	Case adLongVarBinary
		SQLUnwrap = Binary2String(fld.Value)
	Case adLongVarChar
		SQLUnwrap = fld.Value
	Case adLongVarWChar
		SQLUnwrap = fld.Value
	Case adNumeric
		SQLUnwrap = CDbl(fld.Value) ' vbscript doesn't support the DECIMAL type, but can convert it to Double
	'Case adPropVariant
	'    SQLUnwrap = ?
	Case adSingle
		SQLUnwrap = fld.Value
	Case adSmallInt
		SQLUnwrap = fld.Value
	Case adTinyInt
		SQLUnwrap = fld.Value
	Case adUnsignedBigInt
		SQLUnwrap = fld.Value
	Case adUnsignedInt
		SQLUnwrap = fld.Value
	Case adUnsignedSmallInt
		SQLUnwrap = fld.Value
	Case adUnsignedTinyInt
		SQLUnwrap = fld.Value
	'Case adUserDefined
	'    SQLUnwrap = ?
	'Case adVarBinary
	'    SQLUnwrap = ?
	Case adVarChar
		SQLUnwrap = fld.Value
	'Case adVariant
	'    SQLUnwrap = ?
	Case adVarNumeric
		SQLUnwrap = fld.Value
	Case adVarWChar
		SQLUnwrap = fld.Value
	Case adWChar
		SQLUnwrap = fld.Value
	Case Else
		DBFuncs_Err "ModDBFuncs.SQLUnwrap()", "Don't know how to unwrap ADODB.DataTypeEnum: {0}", DataTypeDesc(fld.Type)
	End Select
End Function
Sub DumpRS ( rs )
	Dim i
	%>
	<table border=1>
		<tr>
			<%
			For i = 0 To rs.Fields.Count-1
				%>
				<th>
					<%=rs(i).Name%>
				</th>
				<%
			Next
			%>
		</tr>
		<%
		While Not rs.EOF
			%>
			<tr>
				<%
				For i = 0 To rs.Fields.Count-1
					%>
					<td>
						<%=rs(i).Value%>
					</td>
					<%
				Next
				%>
			</tr>
			<%
			rs.MoveNext
		Wend
		%>
	</table>
	<%
End Sub

Sub DBFuncs_Err ( ByVal sLoc, ByVal sErr, ByVal parameters )
	DBFuncs_Trace sLoc, sErr, parameters
	Err.Raise -1, sLoc, StringFormat ( sErr, parameters )
End Sub

Sub DBFuncs_Trace ( ByVal sLoc, ByVal s, ByVal parameters )
	Dim tmp
	tmp = Trim(Session("DBFUNCS_TRACE")&"")
	If tmp <> "" Then
		s = Server.HtmlEncode(StringFormat(s,parameters))
		Response.Write "<font color=" & tmp & ">DBFuncs_Trace: " & Server.HtmlEncode(sLoc) & ": " & s & "</font><br>" & vbCrLf
	End If
End Sub
%>
