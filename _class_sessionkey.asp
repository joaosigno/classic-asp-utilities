<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Class SessionKey
	Public Key

	Public Default Property Get Value()
		Value = Session(Key)
	End Property

	Public Property Let Value ( ByVal pValue )
		Session(Key) = pValue
	End Property

	Public Property Set Value ( ByRef pValue )
		Set Session(Key) = pValue
	End Property
End Class

Public Function NewSessionKey ( ByVal Key )
	Set NewSessionKey = New SessionKey
	NewSessionKey.Key = Key
End Function
%>