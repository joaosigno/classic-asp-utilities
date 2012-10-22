<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function SelectOption ( ByVal text, ByVal selected )
	If text = selected Then
		selected = " selected"
	Else
		selected = ""
	End If
	%>
		<option<%=selected%>><%=text%>
	<%
End Function

Function SelectOptionValue ( ByVal text, ByVal value, ByVal selected )
	If value = selected Then
		selected = " selected"
	Else
		selected = ""
	End If
	%>
	<option value="<%=value%>"<%=selected%>><%=text%>
	<%
End Function
%>
