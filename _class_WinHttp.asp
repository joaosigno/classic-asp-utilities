<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/

' Simplest Examples:

'	Function DoUrlGet ( ByVal aUrl )
'		With New WinHttp
'			DoUrlGet = .OpenAndSendGet ( aUrl )
'		End With
'	End Function

'	Function DoUrlPost ( ByVal aUrl, ByVal aPostData )
'		With New WinHttp
'			DoUrlPost = .OpenAndSendPost ( aUrl, aPostData )
'		End With
'	End Function

Class WinHttp
	Public Timeout
	Public UserAgent
	Public SslErrorIgnoreFlags
	Public EnableRedirects
	Public EnableHttpsToHttpRedirects
	Public HostOverride
	Public Login
	Public Password

	Public objWinHttp

	Public Sub Class_Initialize()
		Timeout = 59000
		UserAgent = "http_requester/0.1"
		SslErrorIgnoreFlags = 13056 ' 13056: ignore all error, 0: accept no error
		EnableRedirects = True
		EnableHttpsToHttpRedirects = True
		HostOverride = ""
		Login = ""
		Password = ""
		Set objWinHttp = Createobject("WinHttp.WinHttpRequest.5.1")
	End Sub

	Public Sub Open ( ByVal aMethod, ByVal aUrl )
		objWinHttp.SetTimeouts Timeout, Timeout, Timeout, Timeout
		objWinHttp.Open aMethod, aUrl
		If aMethod = "POST" Then
			objWinHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		End If
	End Sub

	Public Function Send ( ByVal aPostData )
		Send = "" ' default to success

		If HostOverride <> "" Then
			objWinHttp.SetRequestHeader "Host", HostOverride
		End If
		objWinHttp.Option(0) = UserAgent
		objWinHttp.Option(4) = SslErrorIgnoreFlags
		objWinHttp.Option(6) = EnableRedirects
		objWinHttp.Option(12) = EnableHttpsToHttpRedirects
		If (Login <> "") And (Password <> "") Then
			objWinHttp.SetCredentials Login, Password, 0
		End If
		On Error Resume Next
		objWinHttp.Send aPostData
		If Err.Number <> 0 Then
			Send = "Error " & Err.Number & " " & Err.Source & " " & Err.Description
		End If
		On Error GoTo 0
	End Function

	Public Function OpenAndSend ( ByVal aMethod, ByVal aUrl, ByVal aPostData )
		Open aMethod, aUrl
		OpenAndSend = Send ( aPostData )
		If "" = OpenAndSend Then
			OpenAndSend = ResultText
		End If
	End Function

	Public Function OpenAndSendGet ( ByVal aUrl )
		OpenAndSendGet = OpenAndSend ( "GET", aUrl, "" )
	End Function

	Public Function OpenAndSendPost ( ByVal aUrl, ByVal aPostData )
		OpenAndSendPost = OpenAndSend ( "POST", aUrl, aPostData )
	End Function

	Public Property Get Status()
		Status = objWinHttp.Status
	End Property
	Public Property Get StatusText()
		StatusText = objWinHttp.StatusText
	End Property
	Public Property Get ResponseText()
		ResponseText = objWinHttp.ResponseText
	End Property

	Public Function ResultText()
		If "200" = Status Then
			ResultText = ResponseText
		Else
			ResultText = "HTTP " & Status & " " & StatusText
		End If
	End Function
End Class ' WinHttp
%>
