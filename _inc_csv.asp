<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function CsvQuote(arg)
	If (Not IsNumeric(arg)) Or (UBound(CsvSplit(arg)) + InStr(arg, vbLf) + InStr(arg, vbCr) > 0) Then
		CsvQuote = """" & Replace(CStr(arg), """", """""") & """"
	Else
		CsvQuote = CStr(arg)
	End If
End Function

Function CsvJoin(ByRef args)
	Dim i, glue
	glue = ""
	If Not IsArray(args) Then args = Array(args)
	For i = LBound(args) To UBound(args)
		If IsArray(args(i)) Then
			CsvJoin = CsvJoin & glue & CsvJoin(args(i))
		Else
			CsvJoin = CsvJoin & glue & CsvQuote(args(i))
		End If
		glue = ","
	Next
End Function

Function p_StrpbrkOrEos(ByRef offset, ByRef sSource, ByVal sFind)
	Dim i, tmp
	p_StrpbrkOrEos = Len(sSource) + 1
	For i = 1 To Len(sFind)
		tmp = InStr(offset, sSource, Mid(sFind, i, 1))
		If tmp > 0 And tmp < p_StrpbrkOrEos Then p_StrpbrkOrEos = tmp
	Next
End Function

Function p_FindEndQuote(ByRef csv, ByVal offset)
	Dim q
	q = offset
	Do
		q = p_StrpbrkOrEos(q, csv, """")
		If Mid(csv, q, 2) = """""" Then ' is the quote followed by another?
			q = q + 2 ' skip this quote and it's match
		Else ' if it's not a double-quote then we're done
			Exit Do
		End If
	Loop
	' found end quote...
	p_FindEndQuote = q
End Function

Function p_CsvUnquoteColumn(ByRef csv, ByRef offset)
	Dim q
	p_CsvUnquoteColumn = ""
	If Mid(csv, offset, 1) = """" Then
		offset = offset + 1 ' skip beginning quote
		q = p_FindEndQuote(csv, offset)
		p_CsvUnquoteColumn = Replace(Mid(csv, offset, q - offset), """""", """")
		offset = q
		If Mid(csv, offset, 1) = """" Then offset = offset + 1
	End If
	If 0 = InStr("," & vbCrLf, Mid(csv, offset, 1)) Then
		q = p_StrpbrkOrEos(offset, csv, "," & vbCrLf)
		p_CsvUnquoteColumn = p_CsvUnquoteColumn & Mid(csv, offset, q - offset)
		offset = q ' update offset
	End If
End Function

' CsvSplitEx() splits a whole file, one line at a time.
' Usage: set offset to 1 for the first call, then pass it back the value it returns each time
Function CsvSplitEx(ByVal csv, ByRef offset)
	Dim ar(), i, count
	ReDim ar(0)
	i = offset
	count = 0
	Do
		ar(count) = p_CsvUnquoteColumn(csv, i)
		count = count + 1
		ReDim Preserve ar(count)
		If Mid(csv, i, 1) <> "," Then Exit Do
		i = i + 1 ' skip comma
	Loop
	' an empty string counts as a single column, so only resize if count is > 1, because array starts off with single element
	If count > 0 Then ReDim Preserve ar(count - 1)
	' skip end-of-line
	If Mid(csv,i,1) = vbCr Then
		If Mid(csv,i+1,1) = vbLf Then
			i = i + 2
		Else
			i = i + 1
		End If
	ElseIf Mid(csv,i,1) = vbLf Then
		If Mid(csv,i+1,1) = vbCr Then
			i = i + 2
		Else
			i = i + 1
		End If
	End If
	offset = i
	CsvSplitEx = ar
End Function

Function CsvSplit(ByVal csv)
	Dim offset
	offset = 1
	CsvSplit = CsvSplitEx(csv, offset)
End Function

Function CsvMultiSplit(ByVal csv)
	Dim ar(), offset, count, last
	ReDim ar(0)
	offset = 1
	last = Len(csv)
	count = 0
	Do
		ar(count) = CsvSplitEx(csv, offset)
		count = count + 1
		ReDim Preserve ar(count)
		If offset > last Then
			Exit Do
		ElseIf InStr(vbCrLf, Mid(csv, offset, 1)) Then
			Select Case Mid(csv, offset, 2)
			Case vbCrLf, vbLf & vbCr
				offset = offset + 2
			Case Else
				offset = offset + 1
			End Select
		Else
			Exit Do
		End If
	Loop
	' an empty string counts as a single row and column, so only resize if count is > 0
	If count > 0 Then ReDim Preserve ar(count - 1)
	CsvMultiSplit = ar
End Function
%>
