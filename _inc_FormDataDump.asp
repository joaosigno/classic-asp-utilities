<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Sub FormDataDump(bolShowOutput, bolEndPageExecution)
  Dim sItem

  'What linebreak character do we need to use?
  Dim strLineBreak
  If bolShowOutput then
    'We are showing the output, so set the line break character
    'to the HTML line breaking character
    strLineBreak = "<br>"
  Else
    'We are nesting the data dump in an HTML comment block, so
    'use the carraige return instead of <br>
    'Also start the HTML comment block
    strLineBreak = vbCrLf
    Response.Write("<!--" & strLineBreak)
  End If
  

  'Display the Request.Form collection
  Response.Write("DISPLAYING REQUEST.FORM COLLECTION" & strLineBreak)
  For Each sItem In Request.Form
    Response.Write(sItem)
    Response.Write(" - [" & Request.Form(sItem) & "]" & strLineBreak)
  Next
  
  
  'Display the Request.QueryString collection
  Response.Write(strLineBreak & strLineBreak)
  Response.Write("DISPLAYING REQUEST.QUERYSTRING COLLECTION" & strLineBreak)
  For Each sItem In Request.QueryString
    Response.Write(sItem)
    Response.Write(" - [" & Request.QueryString(sItem) & "]" & strLineBreak)
  Next

  
  'If we are wanting to hide the output, display the closing
  'HTML comment tag
  If Not bolShowOutput then Response.Write(strLineBreak & "-->")

  'End page execution if needed
  If bolEndPageExecution then Response.End
End Sub
%>