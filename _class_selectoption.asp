<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Class CSelectOption
	Public selected
	Private mCompare

	Public Sub Class_Initialize()
		selected = ""
		mCompare = vbBinaryCompare
	End Sub

	Public Property Get CaseSensitive()
		CaseSensitive = mCompare <> vbBinaryCompare
	End Property
	Public Property Let CaseSensitive ( ByVal Value )
		If Value Then
			mCompare = vbTextCompare
		Else
			mCompare = vbBinaryCompare
		End If
	End Property

	Public Sub Opt ( ByVal text )
		Dim tmp
		If 0 = StrComp ( text, selected, mCompare ) Then
			tmp = " selected"
		Else
			tmp = ""
		End If
		%>
		<option<%=tmp%>><%=text%></option>
		<%
	End Sub

	Public Sub OptVal ( ByVal text, ByVal value )
		Dim tmp
		If 0 = StrComp ( text, selected, mCompare ) Then
			tmp = " selected"
		Else
			tmp = ""
		End If
		%>
		<option value="<%=value%>"<%=tmp%>><%=text%></option>
		<%
	End Sub
End Class
Function NewSelectOption ( ByVal selected )
	Set NewSelectOption = New CSelectOption
	NewSelectOption.selected = selected
End Function
Function NewSelectOptionEx ( ByVal selected, ByVal CaseSensitive )
	Set NewSelectOptionEx = New CSelectOption
	With NewSelectOptionEx
		.selected = selected
		.CaseSensitive = CaseSensitive
	End With
End Function
%>
