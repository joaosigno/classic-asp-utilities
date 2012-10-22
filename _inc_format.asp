<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/

Function StringFormat ( ByVal str, ByVal parameters )
	If Not IsNull(parameters) Then
		' if user accidentally passed a single parameter without putting it in an Array, do it for them
		If Right(TypeName(parameters),2) <> "()" Then parameters = Array(parameters)
		Dim i, tmp
		For i = 0 To UBound(parameters)
			On Error Resume Next
			Err.Clear
			tmp = parameters(i) & ""
			If Err Then
				On Error GoTo 0
				Err.Raise -1, "StringFormat()", "CStr() cannot handle type '" & TypeName(parameters(i)) & "', str=" & str
				tmp = "(" & TypeName(parameters(i)) & ")"
			End If
			On Error GoTo 0
			str = Replace ( str, "{" & i & "}", tmp )
		Next
	End If
	StringFormat = str
End Function

Function SqlEncode ( ByVal s )
	SqlEncode = Replace ( s, "'", "''" )
End Function

Function HtmlFormat ( ByVal s, ByVal parameters )
	'Response.Write "s=" & s & "<br>" & vbCrLf
	If Not IsNull(parameters) Then
		' if user passed a single parameter without putting it in an Array, do it for them
		If Right(TypeName(parameters),2) <> "()" Then parameters = Array(parameters)
		Dim i, tmp
		For i = 0 To UBound(parameters)
			On Error Resume Next
			Err.Clear
			tmp = parameters(i) & ""
			If Err Then
				On Error GoTo 0
				Err.Raise -1, "HtmlFormat()", "cannot handle type '" & TypeName(parameters(i)) & "'"
				tmp = "(" & TypeName(parameters(i)) & ")"
			End If
			On Error GoTo 0
			s = Replace ( s, "{" & i & "}", Server.HtmlEncode(tmp) )
		Next
	End If
	'Response.Write "s=" & s & "<br>" & vbCrLf
	HtmlFormat = s
End Function

Function UrlFormat ( ByVal s, ByVal parameters )
	If Not IsNull(parameters) Then
		' if user accidentally passed a single parameter without putting it in an Array, do it for them
		If Right(TypeName(parameters),2) <> "()" Then parameters = Array(parameters)
		'Response.Write vbCrLf & "<font color=red>s=" & s & ",params=(""" & Join(parameters,""",""") & """)</font><br>" & vbCrLf
		Dim i, tmp
		For i = 0 To UBound(parameters)
			On Error Resume Next
			Err.Clear
			tmp = parameters(i) & ""
			If Err Then
				On Error GoTo 0
				Err.Raise -1, "UrlFormat()", "cannot handle type '" & TypeName(parameters(i)) & "'"
				tmp = "(" & TypeName(parameters(i)) & ")"
			End If
			On Error GoTo 0
			s = Replace ( s, "{" & i & "}", Server.URLEncode(tmp) )
		Next
	'Else
	'	Response.Write vbCrLf & "<font color=red>s=" & s & ",params=null</font><br>" & vbCrLf
	End If
	'Response.Write "s=" & s & "<br>" & vbCrLf
	UrlFormat = s
End Function

Public Function FormatPhone ( phone, sPrefix, bFormat, bBrackets, b10DigitLocal, AreaCodes )
	Dim s, i, c
	s = Trim(phone)
	If (Len(s) > 0) And (Left(s, 1) <> "[") Then
		i = 1
		While i <= Len(s)
			c = Mid(s, i, 1)
			If Asc(c) < Asc("0") Or Asc(c) > Asc("9") Then
				s = Left(s, i - 1) + Right(s, Len(s) - i)
			Else
				i = i + 1
			End If
		Wend
		If Len(s) > 0 Then
			s = sPrefix & s
		End If
		If Len(s) = 10 And bFormat Then
			Dim ar, sOne
			sOne = "1"
			ar = Split(AreaCodes,",")
			For i = LBound(ar) To UBound(ar)
				If Left(s, 3) = ar(i) Then
					sOne = ""
					Exit For
				End If
			Next
			If sOne <> "" Then
				s = sOne & s
			ElseIf Not b10DigitLocal Then
				s = Mid(s, 4)
			End If
		End If
		If bFormat Then
			If Len(s) = 7 Then
				s = Left(s, 3) & " " & Right(s, 4)
			ElseIf Len(s) = 10 Then
				s = Left(s, 3) & " " & Mid(s, 4, 3) & " " & Right(s, 4)
			ElseIf Len(s) = 11 Then
				s = Left(s, 1) & " " & Mid(s, 2, 3) & " " & Mid(s, 5, 3) & " " & Right(s, 4)
			End If
		End If
		If bBrackets And Len(s) Then
			FormatPhone = "[" + s + "]"
		Else
			FormatPhone = s
		End If
	Else
		FormatPhone = phone
	End If
End Function

Function JsEncode ( ByVal s )
	JSEncode = Replace ( s, "'", "\'" )
End Function

Function PathFormat ( s, fmt )
	Dim drive, path, ext, n
	s = Replace(s,"/","\")
	If Mid(s,2,1) = ":" Then
		drive = Left(s,2)
		s = Mid(s,3)
	End If
	n = InStrRev(s,"\")
	If n > 0 Then
		path = Left(s,n)
		s = Mid(s,n+1)
	End If
	n = InStrRev(s,".")
	If n > 0 Then
		ext = Mid(s,n+1)
		s = Left(s,n-1)
	End If
	PathFormat = ""
	While fmt <> ""
		Select Case LCase(Left(fmt,1))
		Case "d" ' drive
			PathFormat = PathFormat & drive
		Case "p" ' path ( non-file
			PathFormat = PathFormat & path
		Case "f" ' file name + extension
			If ext <> "" Then
				PathFormat = PathFormat & s & "." & ext
			Else
				PathFormat = PathFormat & s
			End If
		Case "n" ' file name w/o extension
			PathFormat = PathFormat & s
		Case "x" ' extension
			PathFormat = PathFormat & ext
		Case Else
			PathFormat = PathFormat & Left(fmt,1)
		End Select
		fmt = Mid(fmt,2)
	Wend
End Function
%>
