<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Private base64_state
Private base64_pow2() ' As Long

base64_state = 0
Sub base64_init()
	Dim i, v
	ReDim base64_pow2(8)
	v = 1
	For i = 0 To 8
		base64_pow2(i) = v
		v = v * 2
	Next
	base64_state = 1
End Sub

Class CBase64
	Private Alfabet

	Public Sub Class_Initialize()
		If 0 = base64_state Then base64_init()
		SetStandardEncoding()
	End Sub
	
	Public Sub SetStandardEncoding()
		SetSpecialEncoding "+/"
	End Sub
	
	Public Sub SetFileNameEncoding()
		SetSpecialEncoding "+-"
	End Sub
	
	Public Sub SetSpecialEncoding ( ByVal specials )
		If TypeName(specials) <> "String" Or Len(specials) <> 2 Then
			Err.Raise -1, "Base64.SetSpecialEncoding", "Base64.SetSpecialEncoding expecting String of Len=2"
			Exit Sub
		End If
		Alfabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789" & specials
	End Sub

	Public Function Encode ( ByVal sInput ) ' As String
		Encode = EncodeEx ( sInput, True )
	End Function

	Public Function EncodeEx ( ByVal sInput, ByVal bPad ) ' As String
		Dim x ' As Byte
		Dim v ' As Integer
		Dim topbit ' As Byte
		Dim sOutput ' As String

		Do
			If topbit < 6 Then
				x = x + 1
				v = v * base64_pow2(8)
				If x <= Len(sInput) Then v = v + Asc(Mid(sInput, x, 1))
				topbit = topbit + 8
			End If
			topbit = topbit - 6
			If x > Len(sInput) And 0 = v Then Exit Do
			sOutput = sOutput & Mid(Alfabet, ((v \ base64_pow2(topbit)) And 63) + 1, 1)
			v = v And (base64_pow2(topbit) - 1)
		Loop Until x > Len(sInput) And v = 0
		EncodeEx = sOutput
		If bPad Then
			EncodeEx = EncodeEx & String((8 - ((Len(sOutput)) Mod 4)) Mod 4, "=")
		End If
	End Function

	Public Function Decode ( ByVal sInput ) ' As String
		Dim x ' As Byte
		Dim v ' As Integer
		Dim topbit ' As Byte
		Dim sOutput ' As String
		
		sInput = Replace(sInput, "=", "")
		Do
			Do Until topbit >= 8
				x = x + 1
				v = v * base64_pow2(6)
				If x <= Len(sInput) Then
					v = v + (InStr(Alfabet, Mid(sInput, x, 1)) - 1)
				End If
				topbit = topbit + 6
			Loop
			topbit = topbit - 8
			If x > Len(sInput) And 0 = v Then Exit Do
			sOutput = sOutput & Chr((v \ base64_pow2(topbit)) And 255)
			v = v And (base64_pow2(topbit) - 1)
		Loop Until x > Len(sInput) And v = 0
		Decode = sOutput
	End Function

End Class ' Base64
%>
