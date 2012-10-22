<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Class Vector
	Private m_count
	Private m_data()

	Public Default Property Get Items(ByVal Index) ' Default
		If Index >= 0 And Index < Count Then
			p_Assign Items, m_data(Index)
		End If
	End Property

	Public Property Let Items(ByVal Index, ByVal Value)
		If Index >= Count Then Count = Index + 1
		m_data(Index) = Value
	End Property

	Public Property Set Items(ByVal Index, ByVal Value)
		If Index >= Count Then Count = Index + 1
		Set m_data(Index) = Value
	End Property

	Public Property Get UB()
		UB = UBound(m_data)
	End Property

	Public Property Let UB(ByVal Value)
		If Value < 0 Then Value = 0
		ReDim Preserve m_data(Value)
		If Count > Value Then Count = Value
	End Property

	Public Property Get Count()
		Count = m_count
	End Property

	Public Property Let Count(ByVal Value)
		Dim neededUB
		neededUB = Value
		If neededUB < 0 Then neededUB = 0
		If neededUB > UB Then ' do we need to reallocate?
			UB = Value * 3 \ 2 + 1 ' reallocate with extra capacity
		ElseIf neededUB < UB \ 2 Then ' do we now have too much allocated?
			UB = neededUB ' drop down to requested size to conserve memory
		End If
		m_count = Value
	End Property

	Public Property Get Head()
		If 0 = Count Then
			Err.Raise -1, "Vector.Head.Get()", "Vector is empty"
			Exit Property
		End If
		p_Assign Head, Items(0)
	End Property

	Public Property Let Head(ByVal Value)
		If 0 = Count Then
			Err.Raise -1, "Vector.Head.Let()", "Vector is empty"
			Exit Property
		End If
		Items(0) = Value
	End Property

	Public Property Set Head(ByVal Value)
		If 0 = Count Then
			Err.Raise -1, "Vector.Head.Set()", "Vector is empty"
			Exit Property
		End If
		Set Items(0) = Value
	End Property

	Public Property Get Tail()
		If 0 = Count Then
			Err.Raise -1, "Vector.Tail.Get()", "Vector is empty"
			Exit Property
		End If
		p_Assign Tail, Items(Count - 1)
	End Property

	Public Property Let Tail(ByVal Value)
		If 0 = Count Then
			Err.Raise -1, "Vector.Tail.Let()", "Vector is empty"
			Exit Property
		End If
		Items(Count - 1) = Value
	End Property

	Public Property Set Tail(ByVal Value)
		If 0 = Count Then
			Err.Raise -1, "Vector.Tail.Set()", "Vector is empty"
			Exit Property
		End If
		Set Items(Count - 1) = Value
	End Property

	' vector functions

	Public Sub Insert(ByVal Index, ByVal Value)
		InsertArray Index, Array(Value)
	End Sub

	Public Sub InsertArray(ByVal Index, ByVal ar)
		'p_Trace "InsertArray(" & Index & ",<" & Join(ar, ",") & ">), Count=" & Count
		If Index < 0 Then Index = 0
		If Index >= Count Then Count = Index ' this will cause us to append
		Dim ar_size, i, tmp
		If Not IsArray(ar) Then
			Err.Raise -1, "Vector.InsertArray()", "parameter not an array"
			Exit Sub
		End If
		ar_size = UBound(ar) + 1
		tmp = Count
		Count = tmp + ar_size ' resize array
		'p_Trace "ar_size=" & ar_size & ", FinalIndex=" & Index & ", Count=" & Count & "<br>"
		For i = Count-ar_size-1 To Index Step -1
			'p_Trace "i+ar_size=" & (i+ar_size) & ", i=" & i & "<br>"
			p_Assign m_data(i+ar_size), m_data(i)
		Next
		For i = 0 To ar_size - 1
			p_Assign m_data(Index + i), ar(i)
		Next
	End Sub

	Public Function Remove(ByVal Index)
		Dim ar
		ar = RemoveRange(Index, 1)
		ReDim Preserve ar(0) ' make sure no errors in case we got an empty array
		p_Assign Remove, ar(0)
	End Function

	Public Function RemoveRange(ByVal Index, ByVal NumItems)
		'p_Trace "RemoveRange(Index=" & Index & ",NumItems=" & NumItems & ") Count=" & Count & "<br>"
		If Index < 0 Then NumItems = NumItems + Index: Index = 0  ' if Index negative, reduce NumItems by appropriate number and set Index=0
		If Index + NumItems >= Count Then NumItems = Count - Index ' don't retrieve more items than available
		If NumItems <= 0 Then RemoveRange = Array(): Exit Function  ' no or negative items? return empty array and make no change
		'p_Trace "RemoveRange(): FinalIndex=" & Index & ", FinalNumItems=" & NumItems & ", Count=" & Count & "<br>"
		Dim ar(), i, tmp
		ReDim ar(NumItems - 1)
		For i = 0 To NumItems - 1
			'p_Trace "p_Assign(1) " & i & ", " & (Index+i) & "<br>"
			p_Assign ar(i), m_data(Index + i)
			m_data(Index + i) = Empty
		Next
		tmp = Count

		For i = Index To tmp-NumItems-1
			'p_Trace "p_Assign(2) " & i & ", " & i+NumItems
			p_Assign m_data(i), m_data(i+NumItems)
		Next

		Count = tmp - NumItems
		'p_Trace "Vector.RemoveRange(): ar=<" & Join(ar,",") & ">"
		RemoveRange = ar
	End Function

	Public Function RemoveAll()
		RemoveAll = Count
		Count = 0
	End Function

	Public Function Clear()
		Clear = Count
		Count = 0
	End Function

	' stack (FILO) functions...

	Public Sub Push(ByVal Value)
		Insert Count, Value
	End Sub

	Public Sub PushArray(ByVal ar)
		InsertArray Count, ar
	End Sub

	Public Function Pop()
		p_Assign Pop, Remove(Count - 1)
	End Function

	Public Function PopRange(ByVal NumItems)
		PopRange = RemoveRange(Count - NumItems, NumItems)
	End Function

	' queue (FIFO) functions...

	Public Sub Que(ByVal Value)
		Insert Count, Value
	End Sub

	Public Sub QueArray(ByVal ar)
		InsertArray Count, ar
	End Sub

	Public Function Deque()
		p_Assign Deque, Remove(0)
	End Function

	Public Function DequeRange(ByVal NumItems)
		DequeRange = RemoveRange(0, NumItems)
	End Function

	' advanced function(s)

	Public Function GetArray()
		Dim ar
		ar = m_data
		ReDim Preserve ar(Count - 1)
		GetArray = ar
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
		m_count = 0
		ReDim m_data(0)
	End Sub

	Private Sub Class_Terminate()
		Erase m_data
	End Sub


End Class

Function NewVector ( ByVal ar )
	Set NewVector = New Vector
	NewVector.InsertArray 0, ar
End Function
%>
