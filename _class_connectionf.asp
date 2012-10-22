<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Class ConnectionF
	Public con ' As ADODB.Connection
	
	Public SqlEncoder

	Public Sub Class_Initialize()
		Set SqlEncoder = Nothing
	End Sub

	Public Property Get SqlType()
		SqlType = SqlEncoder.SqlType
	End Property
	Public Property Let SqlType ( ByVal Value )
		Set SqlEncoder = NewSqlEncoder ( Value )
	End Property
	
	Public Function SqlFormat ( ByVal sql, ByVal parameters )
		Err.Source = TypeName(Me) & ".SqlFormat"
		If Not IsArray(parameters) Then parameters = Array(parameters)
		Dim out, i, last, n, match, tmp, s
		last = Len(sql)
		i = 1
		While i <= last
			match = 1
			s = Mid(sql,i,1)
			Select Case s
			Case "'" ' string
				n = InStr(i+1,sql,"'")
				If n > 0 Then
					match = n - i + 1
					s = Mid(sql,i,match)
				End If
			Case "[" ' field name
				n = InStr(i+1,sql,"]")
				If n > 0 Then
					match = n - i + 1
					s = SqlEncoder.Field ( Mid(sql,i+1,match-2) )
				End If
			Case "{"
				n = InStr(i+1,sql,"}")
				If n > 0 Then
					match = n - i + 1
					s = Mid(sql,i,match)
					tmp = Mid(sql,i+1,match-2)
					
					n = Len(tmp)
					Select Case Right(tmp,1)
					Case "d"
						n = CLng(Left(tmp,n-1)) ' n is now index into parameters
						s = SqlEncoder.DateOnly ( parameters(n) )
					Case "t"
						If Right(tmp,2) = "dt" Then
							n = CLng(Left(tmp,n-2)) ' n is now index into parameters
							s = SqlEncoder.DateTime ( parameters(n) )
						Else
							n = CLng(Left(tmp,n-1)) ' n is now index into parameters
							s = SqlEncoder.TimeOnly ( parameters(n) )
						End If
					Case "s"
						tmp = Left(tmp,n-1)
						
						n = Replace(tmp,"%","") ' n is now index into parameters
						If IsNumeric(n) Then ' otherwise it's an invalid parameter index
							n = CLng(n)
							If n >= 0 And n <= UBound(parameters) Then
								s = SqlEncoder.String ( Replace ( tmp, CStr(n), parameters(n) & "" ) )
							End If
						End If
					Case "_"
						n = CLng(Left(tmp,n-1)) ' n is now index into parameters
						s = SqlEncoder.Field ( parameters(n) )
					Case "?"
						n = CLng(Left(tmp,n-1)) ' n is now index into parameters
						s = SqlEncoder.Guess ( parameters(n) )
					Case Else
						If IsNumeric(Right(tmp,1)) Then
							n = CLng(tmp) ' n is now index into parameters
							' technically this works, but logically it's wrong.
							' should we have a SqlEncoder.Literal?
							s = SqlEncoder.Numeric ( parameters(n) )
						Else
							Err.Raise -1,, "Unsupported suffix: " & Right(tmp,1)
						End If
					End Select
				End If
			'Case Else
				' TODO: search for next special-case...
			End Select
			out = out & s
			i = i + match
		Wend
		SqlFormat = out
	End Function

	Public Function SqlFormatTop(ByVal nTop, ByVal sql, ByVal parameters)
		Err.Source = TypeName(Me) & ".SqlFormatTop"
		Select Case LCase(Left(sql, 7))
		Case "select ", "update "
			sql = SqlFormat(sql, parameters)
			If SqlEncoder.HasLimit Then
				sql = sql & " LIMIT 0," & (nTop - 1)
			ElseIf SqlEncoder.HasTop Then
				sql = Left(sql, 7) & " top " & nTop & " " & Mid(sql, 8)
			Else
				Err.Raise -1,, "SqlEncoder doesn't support TOP or LIMIT"
			End If
		Case Else
			Err.Raise -1,, "must be called with a select or update query"
			Exit Function
		End Select
		SqlFormatTop = sql
	End Function

	Public Function Execute(ByVal sql, ByVal parameters) ' As ADODB.Recordset
		Dim tmp
		tmp = SqlFormat(sql, parameters)
		Set Execute = con.Execute(tmp)
	End Function

	Public Function ExecuteTop(ByVal nTop, ByVal sql, ByVal parameters) ' As ADODB.Recordset
		Dim tmp
		tmp = SqlFormatTop(nTop, sql, parameters)
		Set ExecuteTop = con.Execute(tmp)
	End Function

	Public Sub Close()
		con.Close
	End Sub

	Public Sub Class_Terminate()
		Set con = Nothing
	End Sub
End Class

Function NewConnectionF ( ByRef con, ByVal SqlType )
	Set NewConnectionF = New ConnectionF
	With NewConnectionF
		Set .con = con
		.SqlType = SqlType
	End With
End Function
%>
