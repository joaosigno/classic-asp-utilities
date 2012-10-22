<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Class Serializer
	Public d, bSave

	Public Sub Class_Initialize()
		Set d = CreateObject("Scripting.Dictionary")
		bSave = Empty
	End Sub

	Function Serialize ( ByVal aKey, ByRef aValue )
		If IsEmpty(bSave) Then Err.Raise -1, "Serializer", "bSave was not set" : Exit Function

		If bSave Then
			d(aKey) = aValue
		Else
			aValue = d(aKey)
		End If

		Serialize = aValue
	End Function
End Class

Function NewSerializer ( ByVal bSave )
	Set NewSerializer = New Serializer
	NewSerializer.bSave = bSave
End Function
%>