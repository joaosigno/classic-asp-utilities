<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Sub Paginator_Trace ( ByVal s )
	If Session("PAGINATOR_TRACE") Then
		Response.Write HtmlFormat("<font color=red>{0}</font>" & vbCrLf,s)
	End If
End Sub

Class Paginator
	Private m_PageSize, m_PageNumber, m_rs, m_RelativePosition

	Public Sub Init ( ByVal PageSize, ByVal PageNumber, ByRef con, ByVal sql )
		m_PageSize = CLng(PageSize)
		m_PageNumber = CLng(PageNumber)
		Set m_rs = NewRSReadOnly(con)
		m_rs.CursorLocation = adUseClient
		m_rs.CacheSize = m_PageSize
		m_rs.Open sql
		If m_rs.RecordCount > 0 Then
			MoveFirst
		End If
	End Sub

	Public Property Get AbsoluteCount()
		AbsoluteCount = m_rs.RecordCount
	End Property

	Public Property Get RelativeCount()
		RelativeCount = LastAbsolute - FirstAbsolute + 1
	End Property
	
	Public Default Property Get Fields ( ByVal Index )
		Set Fields = m_rs.Fields(Index)
	End Property

	Public Property Get PageSize()
		PageSize = m_PageSize
	End Property

	Public Property Get PageNumber()
		PageNumber = m_PageNumber
	End Property

	Public Property Get AbsolutePosition()
		AbsolutePosition = CLng(m_rs.AbsolutePosition)
	End Property
	
	Public Property Get RelativePosition()
		RelativePosition = m_RelativePosition
	End Property
	
	Private Function CalcAbsolute ( ByVal Position )
		CalcAbsolute = (m_PageSize * (m_PageNumber-1)) + CLng(Position)
	End Function
	
	Public Property Let RelativePosition ( ByVal value )
		m_RelativePosition = CLng(value)
		If m_RelativePosition < 0 Then
			m_RelativePosition = 0
		ElseIf m_RelativePosition > m_PageSize+1 Then
			m_RelativePosition = m_PageSize+1
		End If
		m_rs.AbsolutePosition = CalcAbsolute(m_RelativePosition)
	End Property
	
	Public Property Get BOF()
		BOF = (m_RelativePosition < 1) Or m_rs.BOF
	End Property
	
	Public Property Get EOF()
		EOF = (m_RelativePosition > m_PageSize) Or m_rs.EOF
	End Property
	
	Public Sub Move ( ByVal NumberRecords )
		RelativePosition = RelativePosition + NumberRecords
	End Sub
	
	Public Sub MoveFirst()
		RelativePosition = 1
	End Sub
	
	Public Sub MoveLast()
		RelativePosition = RelativeCount
	End Sub
	
	Public Sub MoveNext()
		Move 1
	End Sub
	
	Public Sub MovePrevious()
		Move -1
	End Sub
	
	Public Property Get FirstAbsolute()
		FirstAbsolute = CalcAbsolute(1)
	End Property
	
	Public Property Get LastAbsolute()
		LastAbsolute = CalcAbsolute(m_PageSize)
		If LastAbsolute > m_rs.RecordCount Then LastAbsolute = CLng(m_rs.RecordCount)
	End Property
	
	Public Property Get LastPage()
		LastPage = CLng((m_rs.RecordCount-1) \ PageSize + 1)
	End Property
	
	Public Property Get Recordset()
		Set Recordset = m_rs
	End Property

	Public Sub NavLinks ( ByVal NumPages, ByVal QueryPage, ByVal ExtraQueryParameters )
		Dim ar, url, i
		url = "?page={0}"
		If IsArray(ExtraQueryParameters) Then
			ReDim ar(UBound(ExtraQueryParameters)\2+1)
			For i = 0 To UBound(ar)-1
				url = url & "&" & ExtraQueryParameters(i*2) & "={" & (i+1) & "}"
				ar(i+1) = ExtraQueryParameters(i*2+1)
			Next
		Else
			ReDim ar(0)
		End If
		Paginator_Trace "url=" & url & ",params=" & Join(ar,",")
		Dim first, last
		first = m_PageNumber - NumPages \ 2
		If first < 1 Then first = 1
		last = first + NumPages - 1
		If last > LastPage Then last = LastPage
		' recalc first one more time, in case we're near the last page...
		first = last - NumPages + 1
		If first < 1 Then first = 1
		If False Then ' Debug
			Response.Write "LastAbsolute=" & LastAbsolute & ",LastPage=" & LastPage & "<br>" & vbCrLf
			Response.Write "page=" & m_PageNumber & ",first=" & first & ",last=" & last & "<br>" & vbCrLf
		End If
		If m_PageNumber = 1 Then
			%>
			&nbsp;First
			&nbsp;Prev
			<%
		Else
			ar(0) = 1
			%>
			&nbsp;<a href="<%=UrlFormat(url,ar)%>">First</a>
			<% ar(0) = m_PageNumber-1 %>
			&nbsp;<a href="<%=UrlFormat(url,ar)%>">Prev</a>
			<%
		End If
		For i = first To last
			If CLng(i) = CLng(m_PageNumber) Then
				%>
				&nbsp;<%=i%>
				<%
			Else
				ar(0) = i
				%>
				&nbsp;<a href="<%=UrlFormat(url,ar)%>"><%=i%></a>
				<%
			End If
		Next
		If m_PageNumber = last Then
			%>
			&nbsp;Next
			&nbsp;Last
			<%
		Else
			ar(0) = m_PageNumber + 1
			%>
			&nbsp;<a href="<%=UrlFormat(url,ar)%>">Next</a>
			<% ar(0) = LastPage %>
			&nbsp;<a href="<%=UrlFormat(url,ar)%>">Last</a>
			<%
		End If
		%>
		&nbsp;
		<%
	End Sub
End Class

Public Function NewPaginator ( ByVal PageSize, ByVal PageNumber, ByRef con, ByVal sql )
	Dim c
	Set c = New Paginator
	c.Init PageSize, PageNumber, con, sql
	Set NewPaginator = c
End Function
%>