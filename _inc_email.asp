<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Sub SendEmail (     sFromEmail, sToEmail, sBccEmail, sSubject,         sBody, sSMTPServer, sSMTPUser, sSMTPPass )
	SendEmailEx     sFromEmail, sToEmail, sBccEmail, sSubject, False,  sBody, sSmtpServer, sSmtpUser, sSmtpPass
End Sub

Sub SendHtmlEmail ( sFromEmail, sToEmail, sBccEmail, sSubject,         sBody, sSmtpServer, sSmtpUser, sSmtpPass )
	SendEmailEx     sFromEmail, sToEmail, sBccEmail, sSubject, True,   sBody, sSmtpServer, sSmtpUser, sSmtpPass
End Sub

Sub SendEmailEx (   sFromEmail, sToEmail, sBccEmail, sSubject, IsHtml, sBody, sSMTPServer, sSMTPUser, sSMTPPass )
	Dim msg

	Dim cfg
	Set cfg = CreateObject("CDO.Configuration")
	Const schemapath = "http://schemas.microsoft.com/cdo/configuration/"
	cfg.Fields.Item(schemapath & "sendusing") = 2 ' cdoSendUsingPort
	If sSMTPServer <> "" Then
		cfg.Fields.Item(schemapath & "smtpserver") = sSMTPServer
	End If
	cfg.Fields.Item(schemapath & "smtpserverport") = 25
	If sSMTPUser <> "" Or sSMTPPass <> "" Then
		cfg.Fields.Item(schemapath & "smtpauthenticate") = 1
		cfg.Fields.Item(schemapath & "sendusername") = sSMTPUser
		cfg.Fields.Item(schemapath & "sendpassword") = sSMTPPass
	End If
	cfg.Fields.Update()
	Set msg = CreateObject("CDO.Message")
	Set msg.Configuration = cfg
	If sFromEmail = "" Then
		msg.From = "Anonymous"
	Else
		msg.From = sFromEmail
	End If

	msg.To = SplitEmailList ( sToEmail )
	msg.Bcc = SplitEmailList ( sBccEmail )
	'msg.Importance = 1


	' if you want to add an attachment...
	' uncomment the next line
	' msg.AttachFile ( "c://autoexec.bat" )

	If IsHtml Then
		'msg.BodyFormat = 0
		'msg.MailFormat = 0
		msg.HtmlBody = sBody
	Else
		msg.TextBody = sBody
	End If
	msg.Subject = sSubject

	' send it
	msg.Send()

	' release object
	Set msg = Nothing
End Sub

Function SplitEmailList ( ByVal Emails )
	Dim sEmailList, nEmail, sMail
	Emails = Replace ( Emails, " ", ";" )
	Emails = Replace ( Emails, ",", ";" )
	sEmailList = Split ( Emails, ";" )
	sMail = ""

	For nEmail = 0 To UBound(sEmailList)
		If sEmailList(nEmail) <> "" Then
			sMail = sMail & sEmailList(nEmail) & ";"
		End If
	Next
	SplitEmailList = sMail
End Function
%>
