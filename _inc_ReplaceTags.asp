<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/

' The ReplaceTags() function searchs for and replaces the
' occurance of tags within the embedded string with contents
' that you specify. The callback object "cb" is used to give you
' control of how to identify tags and what to replace them with
' 
' The following example shows how to replace <b> with <strong> in html markup:
'
' Class ReplaceBWithStrong
'   Public Function Parse ( ByRef src, ByVal Index, ByRef TagClose )
'     Parse = 0 ' default == not a tag
'     If Mid(src,Index,4) = "<!--" Then
'       Parse = 4 ' opening tag length = 4
'       TagClose = "-->"
'       Exit Function
'     End If
'     If Mid(src,Index,1) = "<" Then
'       Parse = 1 ' opening tag length = 1
'       TagClose = ">"
'     End If
'   End Function
'   Public Function ReplaceTag ( ByVal TagOpen, ByVal TagBody, ByVal TagClose )
'     If Trim(LCase(tag)) = "b" Then
'       ReplaceTag = "<strong>"
'     Else
'       ReplaceTag = TagOpen & TagBody & TagClose
'     End If
'   End Function
' End Class
Public Function ReplaceTags ( ByVal str, ByRef cb )
	ReplaceTags = ReplaceTagsEx ( str, cb, vbBinaryCompare )
End Function
Public Function ReplaceTagsEx ( ByVal str, ByRef cb, ByVal compare ) ' see above for cb specification
	Err.Source = "ReplaceTags"
	Dim out, i, TagLen, TagOpen, TagClose, last, n, tmp, IsTag
	last = Len(str)
	i = 1
	While i <= last
		TagClose = "" ' try to avoid confusing our users
		TagLen = cb.Parse ( str, i, TagClose )
		If TagLen > 0 Then
			TagOpen = Mid ( str, i, TagLen )
			i = i + TagLen
			n = InStr ( i, str, TagClose, compare )
			If 0 = n Then
				n = last
			Else
				' get the actual value of TagClose in the source string ( case might be different )
				TagClose = Mid ( str, n, Len(TagClose) )
			End If
			IsTag = True
			tmp = cb.ReplaceTag ( TagOpen, Mid ( str, i, n-i ), TagClose, IsTag )
			If IsTag Then
				i = n + Len(TagClose)
			Else
				i = i - TagLen ' back-track
				tmp = Mid ( str, i, 1 )
				i = i + 1
			End If
		Else
			tmp = Mid ( str, i, 1 )
			i = i + 1
		End If
		out = out & tmp
	Wend
	ReplaceTagsEx = out
End Function
%>
