<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function ValidateEmail ( ByVal eml )
	If UBound(Split(eml,"@")) <> 1 Then
		ValidateEmail = "The email address '" & eml & "' must have one and only one '@'"
	ElseIf 0 = InStr(eml,".") Then
		ValidateEmail = "The email address '" & eml & "' must have a period"
	ElseIf InStrRev(eml,".") < InStr(eml,"@") Then
		ValidateEmail = "The email address '" & eml & "' is missing a '.' after the '@'"
	ElseIf 0 <> InStr("@.",Left(eml,1)) Then
		ValidateEmail = "email address '" & eml & "' must not begin with . or @"
	ElseIf 0 <> InStr("@.",Right(eml,1)) Then
		ValidateEmail = "email address '" & eml & "' must not end with . or @"
	ElseIf 0 <> InStr(eml,"@.") Or 0 <> InStr(eml,".@") Then
		ValidateEmail = "email address '" & eml & "' must not have a . immediately before or after the @"
	ElseIf 0 <> InStr(eml,"..") Then
		ValidateEmail = "email address '" & eml & "' must not have two .. right together"
	ElseIf 0 <> InStr(eml,";") Then
		ValidateEmail = "email address '" & eml & "' must not have a ;"
	Else
		Dim re, i, matches
		Set re = New RegExp
		re.IgnoreCase = True
		re.Global = True
		re.Pattern = "[^A-Za-z0-9_@.]"
		Set matches = re.Execute(eml)
		If matches.Count Then
			ValidateEmail = "email address '" & eml & "' has " & matches.Count & " illegal character(s), the first of which is: " & Replace(matches(0).Value," ","(space)")
		Else
			ValidateEmail = ""
		End If
	End If
End Function

Function ValidateEmailEx ( ByRef emls )
	ValidateEmailEx = ""

	' now check each individual email for validity
	Dim ar, i
	ar = Split(Replace(emls,",",";"),";")
	For i = 0 To UBound(ar)
		If Trim(ar(i)) <> "" Then
			ValidateEmailEx = ValidateEmail(ar(i))
			If ValidateEmailEx <> "" Then Exit Function ' one failure is good enough
		End If
	Next
End Function
%>
