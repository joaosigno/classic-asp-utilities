<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Class Matrix
	Private m_rows, m_cols
	Private m_data()

	Private Function m_Valid ( ByVal Row, ByVal Col, ByVal Src ) ' As Boolean
		m_Valid = False
		Src = "Matrix." & Src & "()"
		If Row < 0 Or Row >= m_rows Then
			Err.Raise -1, Src, Src & " called with invalid Row (" & Row & "), Rows=" & m_rows
		ElseIf Col < 0 Or Col >= m_cols Then
			Err.Raise -1, Src, Src & " called with invalid Col (" & Col & "), Cols=" & m_cols
		Else
			m_Valid = True
		End If
	End Function

	Public Default Property Get Items(ByVal Row, ByVal Col) ' Default
		'If m_Valid(Row,Col,"Items.Get") Then
			p_Assign Items, m_data(Row,Col)
		'End If
	End Property

	Public Property Let Items(ByVal Row, ByVal Col, ByVal Value)
		'If m_Valid(Row,Col,"Items.Let") Then
			m_data(Row,Col) = Value
		'End If
	End Property

	Public Property Set Items(ByVal Index, ByVal Value)
		'If m_Valid(Row,Col,"Items.Set") Then
			Set m_data(Row,Col) = Value
		'End If
	End Property

	Public Property Get Rows()
		Rows = m_rows
	End Property

	Public Property Let Rows(ByVal Value)
		Resize Value, m_cols
	End Property

	Public Property Get Cols()
		Cols = m_cols
	End Property

	Public Property Let Cols(ByVal Value)
		Resize Rows, Value
	End Property

	Public Sub Resize ( ByVal NewRows, ByVal NewCols )
		Dim tmp(), ur, uc, i, j
		tmp = m_data
		ReDim m_data ( NewRows-1, NewCols-1 )
		If NewRows < m_rows Then
			ur = NewRows
		Else
			ur = m_rows
		End If
		If NewCols < m_cols Then
			uc = NewCols
		Else
			uc = m_cols
		End If
		For j = 0 To ur-1
			For i = 0 To uc-1
				m_data(j,i) = tmp(j,i)
			Next
		Next
	End Function

	' private function(s)

	Private Sub p_Trace(ByVal s)
		'Response.Write "<font color=red>" & Server.HtmlEncode(s) & "</font><br>"
	End Sub

	Private Sub p_Assign(ByRef dst, ByVal Value)
		If IsObject(Value) Then
			Set dst = Value
		Else
			dst = Value
		End If
	End Sub

	Private Sub Class_Initialize()
		Resize 1, 1
	End Sub

	Private Sub Class_Terminate()
		Erase m_data
	End Sub

End Class
%>
