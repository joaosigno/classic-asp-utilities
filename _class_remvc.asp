<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/

Class ReMvc
	Public model, view, controller
	Public Sub Class_Initialize()
		Set model = Nothing
		Set view = Nothing
		Set controller = Nothing
	End Sub
	Public Sub SetMvc ( ByRef pModel, ByRef pView, ByRef pController )
		Set model = pModel
		Set view = pView
		Set controller = pController
		If Not model Is Nothing Then Set model.mvc = Me
		If Not view Is Nothing Then Set view.mvc = Me
		If Not controller Is Nothing Then Set controller.mvc = Me
	End Sub
	Public Property Get Method()
		Method = UCase(Trim(Request.ServerVariables("REQUEST_METHOD")))
	End Property
	Public Property Get IsGet()
		IsGet = ("GET" = Method)
	End Property
	Public Property Get IsPost()
		IsPost = ("POST" = Method)
	End Property
	Public Sub Trace ( ByVal msg )
		Response.Write msg & "<br>"
	End Sub
End Class ' ReMvc
Public Function NewReMvc ( ByRef pModel, ByRef pView, ByRef pController )
	Set NewReMvc = New ReMvc
	NewReMvc.SetMvc pModel, pView, pController
	pController.Page_Load
End Function ' NewReMvc

Class ReMvcVar
	Public Name
	Private mValue

	Public Default Property Get Value()
		Value = mValue
	End Property
	Public Property Let Value ( ByVal pValue )
		mValue = pValue
	End Property

	Public Property Get Checked()
		Checked = Trim(mValue) <> ""
	End Property
	Public Property Let Checked ( ByVal pValue )
		If pValue Then
			mValue = "checked"
		Else
			mValue = ""
		End If
	End Property

	Public Sub Class_Initialize()
		Set mAttributes = Nothing
	End Sub

	Public Sub Init ( ByVal pName, ByVal pValue, ByVal pDefault )
		Name = pName
		Value = pValue
		If Trim(Value) = "" Then Value = pDefault
	End Sub
	Public Sub InitSession ( ByVal pName, ByVal pValue, ByVal pDefault )
		Name = pName
		Value = pValue
		If Trim(Value) = "" Then Value = Session(pName)
		If Trim(Value) = "" Then Value = pDefault
	End Sub
	Public Sub InitTrim ( ByVal pName, ByVal pValue, ByVal pDefault )
		Name = pName
		Value = Trim(pValue)
		If Trim(Value) = "" Then Value = Trim(pDefault)
	End Sub
	Public Sub InitSessionTrim ( ByVal pName, ByVal pValue, ByVal pDefault )
		Name = pName
		Value = Trim(pValue)
		If Trim(Value) = "" Then Value = Trim(Session(pName))
		If Trim(Value) = "" Then Value = Trim(pDefault)
	End SUb

	Public Function ToUrl()
		ToUrl = Server.URLEncode(Name) & "=" & Server.URLEncode(Value)
	End Function

	Public Sub Att ( ByVal pName, ByVal pValue )
		If mAttributes Is Nothing Then Set mAttributes = CreateObject("Scripting.Dictionary")
		mAttributes(LCase(Trim(pName))) = pValue
	End Sub

	Public Sub HtmlInput ( ByVal pType )
		Response.Write "<input name='"&Server.HtmlEncode(Name)&"' value='"&Server.HtmlEncode(Value&"")&"' type='"&Server.HtmlEncode(pType)&"'"
		mWriteAttributes
		Response.Write "/>"
	End Sub
	Public Sub HtmlSubmit ( ByVal aValue )
		Value = aValue
		HtmlInput "submit"
	End Sub

	Public Sub HtmlSelect ( ByRef Options )
		Dim key, sel, opt
		Response.Write vbCrLf & "<select name='"&Server.HtmlEncode(Name)&"'"
		mWriteAttributes
		Response.Write ">" & vbCrLf
		sel = CStr(Value & "")
		If IsArray(Options) Then
			For key = 0 To UBound(Options)
				opt = CStr(Options(key) & "")
				Response.Write "<option"
				If opt = sel Then Response.Write " selected"
				Response.Write ">" & Server.HtmlEncode(opt) & "</option>" & vbCrLf
			Next
		Else ' must be a dictionary...
			For Each Key In Options.Keys
				opt = CStr(key & "")
				Response.Write "<option value='" & Server.HtmlEncode(opt) & "'"
				If opt = sel Then Response.Write " selected"
				Response.Write ">" & Server.HtmlEncode(Options(Key) & "") & "</option>" & vbCrLf
			Next
		End If
		Response.Write "</select>"
	End Sub

	Public Sub HtmlCheckbox()
		Response.Write "<input type='checkbox' name='"&Server.HtmlEncode(Name)&"' value='checked'"
		If (Value & "") <> "" Then Response.Write " checked"
		mWriteAttributes
		Response.Write "/>" & vbCrLf
	End Sub

	Public Sub HtmlTextArea()
		Response.Write "<textarea name='"&Server.HtmlEncode(Name)&"'"
		mWriteAttributes
		Response.Write ">" & Value & "</textarea>"
	End Sub

	Private Sub mWriteAttributes()
		If mAttributes Is Nothing Then Exit Sub
		Dim Key
		If Not mAttributes.Exists("id") Then Response.Write " id='" & Server.HtmlEncode(Name) & "'"
		For Each Key In mAttributes.Keys
			Response.Write " " & Server.HtmlEncode(Key) & "='" & Server.HtmlEncode(mAttributes(Key)) & "'"
		Next
	End Sub

	Private mAttributes
End Class ' ReMvcVar

Public Function NewReMvcVarQuery ( ByVal pName, ByVal pDefault )
	Set NewReMvcVarQuery = New ReMvcVar
	NewReMvcVarQuery.Init pName, Request.QueryString(pName), pDefault
End Function
Public Function NewReMvcVarQueryTrim ( ByVal pName, ByVal pDefault )
	Set NewReMvcVarQueryTrim = New ReMvcVar
	NewReMvcVarQueryTrim.InitTrim pName, Request.QueryString(pName), pDefault
End Function

Public Function NewReMvcVarForm ( ByVal pName, ByVal pDefault )
	Set NewReMvcVarForm = New ReMvcVar
	NewReMvcVarForm.Init pName, Request.Form(pName), pDefault
End Function
Public Function NewReMvcVarFormTrim ( ByVal pName, ByVal pDefault )
	Set NewReMvcVarFormTrim = New ReMvcVar
	NewReMvcVarFormTrim.InitTrim pName, Request.Form(pName), pDefault
End Function

Public Function NewReMvcVar ( ByVal pName, ByVal pDefault )
	Set NewReMvcVar = New ReMvcVar
	NewReMvcVar.Init pName, Request(pName), pDefault
End Function ' NewReMvcVar
Public Function NewReMvcVarTrim ( ByVal pName, ByVal pDefault )
	Set NewReMvcVarTrim = New ReMvcVar
	NewReMvcVarTrim.InitTrim pName, Request(pName), pDefault
End Function ' NewReMvcVar

' session variants:

Public Function NewReMvcVarQuerySession ( ByVal pName, ByVal pDefault )
	Set NewReMvcVarQuerySession = New ReMvcVar
	NewReMvcVarQuerySession.InitSession pName, Request.QueryString(pName), pDefault
End Function
Public Function NewReMvcVarQuerySessionTrim ( ByVal pName, ByVal pDefault )
	Set NewReMvcVarQuerySessionTrim = New ReMvcVar
	NewReMvcVarQuerySessionTrim.InitSessionTrim pName, Request.QueryString(pName), pDefault
End Function
%>
