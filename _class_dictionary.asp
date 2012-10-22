<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Class Dictionary
	' internally, the top element is never used, so UBound actually = Count
	Private m_debug
	Private m_count
	Private m_keys()
	Private m_data()
	Private m_default

	Public Default Property Get Items(ByVal IndexKey)
		Dim Index
		If Not IsNumeric(IndexKey) Then
			Index = Find(IndexKey)
			If Err Then Exit Property
		Else
			On Error Resume Next
			Err.Clear
			Index = CLng(IndexKey)
			If Err Then
				On Error GoTo 0
				Err.Raise -1, "Dictionary{" & m_debug & "}.Items.Get(" & IndexKey & ")", "Invalid Index: " & IndexKey
				Exit Property
			End If
		End If
		If Index < 0 Or Index >= Count Then
			If Not IsEmpty(m_default) Then
				p_Assign Items, m_default
				Exit Property
			End If
			Err.Raise -1, "Dictionary{" & m_debug & "}.Items.Get(" & IndexKey & ")", "Bad Key Or Index (" & IndexKey & ") And No Default Available: Index=" & Index & ", Count=" & Count
			'Items = "??? Bad Key or Index"
			Exit Property
		End If
		p_Assign Items, m_data(Index)
	End Property

	Public Property Let Items(ByVal IndexKey, ByVal Value)
		'p_Trace "Items.Let(" & IndexKey & ")"
		Dim Index
		If Not IsNumeric(IndexKey) Then
			Index = Find(IndexKey)
			If Err Then Exit Property
			If Index = -1 Then
				'p_Trace "Items.Let(" & IndexKey & ") - calling Add()"
				Add IndexKey, Value
				Exit Property
			End If
		Else
			On Error Resume Next
			Err.Clear
			Index = CLng(IndexKey)
			If Err Then
				On Error GoTo 0
				Err.Raise -1, "Dictionary{" & m_debug & "}.Items.Let(" & IndexKey & ")", "Invalid Index: " & IndexKey
				Exit Property
			End If
		End If
		If Index < 0 Or Index >= Count Then
			Err.Raise -1, "Dictionary{" & m_debug & "}.Items.Let(" & IndexKey & ")", "Invalid Index: " & Index & ", Count=" & Count
			Exit Property
		End If
		m_data(Index) = Value
	End Property

	Public Property Set Items(ByVal IndexKey, ByVal Value)
		'p_Trace "Items.Set(" & IndexKey & ")"
		Dim Index
		If Not IsNumeric(IndexKey) Then
			Index = Find(IndexKey)
			If Err Then Exit Property
			If Index = -1 Then
				'p_Trace "Items.Set(" & IndexKey & ") - calling Add()"
				Add IndexKey, Value
				Exit Property
			End If
		Else
			On Error Resume Next
			Err.Clear
			Index = CLng(IndexKey)
			If Err Then
				On Error GoTo 0
				Err.Raise -1, "Dictionary{" & m_debug & "}.Items.Set(" & IndexKey & ")", "Invalid Index: " & IndexKey
				Exit Property
			End If
		End If
		If Index < 0 Or Index >= Count Then
			Err.Raise -1, "Dictionary{" & m_debug & "}.Items.Set(" & IndexKey & ")", "Invalid Index: " & Index & ", Count=" & Count
			Exit Property
		End If
		Set m_data(Index) = Value
	End Property

	Public Property Get ByKey(ByVal Key)
		Dim Index
		Index = Find(Key)
		If -1 = Index Then
			If Not IsEmpty(m_default) Then
				p_Assign ByKey, m_default
				Exit Property
			End If
			Err.Raise -1, "Dictionary{" & m_debug & "}.ByKey.Get(" & Key & ")", "Bad Key (" & Key & ") And No Default Available: Index=" & Index & ", Count=" & Count
			Exit Property
		Else
			p_Assign ByKey, m_data(Index)
		End If
	End Property

	Public Property Let ByKey(ByVal Key, ByVal Value)
		Dim Index
		Index = Find(Key)
		If -1 = Index Then
			Add Key, Value
			Exit Property
		Else
			m_data(Index) = Value
		End If
	End Property

	Public Property Set ByKey(ByVal Key, ByVal Value)
		Dim Index
		Index = Find(Key)
		If -1 = Index Then
			Add Key, Value
			Exit Property
		Else
			Set m_data(Index) = Value
		End If
	End Property

	Public Property Get ByIndex(ByVal Index)
		If Index < 0 Or Index >= Count Then
			If Not IsEmpty(m_default) Then
				p_Assign ByIndex, m_default
				Exit Property
			End If
			Err.Raise -1, "Dictionary{" & m_debug & "}.ByIndex.Get(" & Index & ")", "Bad Index (" & Index & ") And No Default Available: Count=" & Count
			Exit Property
		Else
			p_Assign ByIndex, m_data(Index)
		End If
	End Property

	Public Property Let ByIndex(ByVal Index, ByVal Value)
		If Index < 0 Or Index >= Count Then
			Err.Raise -1, "Dictionary{" & m_debug & "}.ByIndex.Let(" & Index & ")", "Invalid Index: " & Index & ", Count=" & Count
		Else
			m_data(Index) = Value
		End If
	End Property

	Public Property Set ByIndex(ByVal Index, ByVal Value)
		If Index < 0 Or Index >= Count Then
			Err.Raise -1, "Dictionary{" & m_debug & "}.ByIndex.Set(" & Index & ")", "Invalid Index: " & Index & ", Count=" & Count
		Else
			Set m_data(Index) = Value
		End If
	End Property

	Public Property Get UB()
		UB = UBound(m_keys)
	End Property

	Private Property Let UB(ByVal Value)
		Dim minUB
		minUB = m_count
		If Value < minUB Then Value = minUB ' don't allow this to be a shortcut to kill the end of the list
		If Value < 0 Then Value = 0 ' don't allow crash
		ReDim Preserve m_keys(Value)
		ReDim Preserve m_data(Value)
	End Property

	Public Property Get Count()
		Count = m_count
	End Property

	' this MUST be Private, else you can end up with an unsorted array or key conflicts
	Private Property Let Count(ByVal Value)
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

	Public Property Get Default()
		p_Assign Default, m_default
	End Property

	Public Property Let Default(ByVal Value)
		m_default = Value
	End Property

	Public Property Set Default(ByVal Value)
		Set m_default = Value
	End Property

	Public Function FindEx(ByVal Key, ByVal bExact)
		' if bExact is False, this function returns insertion index
		' this function performs a binary search to find the index, very fast
		If TypeName(Key)="Null" Then Key = ""
		If IsObject(Key) Then
			FindEx = -1
			Err.Raise -1, "Dictionary{" & m_debug & "}.FindEx([" & TypeName(Key) & "])", "Keys cannot be objects"
			Exit Function
		End If
		'If IsNumeric(Key) Then
		'	FindEx = -1
		'	Err.Raise -1, "Dictionary{" & m_debug & "}.FindEx(" & Key & ")", "Keys cannot be numeric"
		'	Exit Function
		'End If
		Dim lo, hi
		lo = 0
		hi = Count
		'p_Trace "FindEx() called with key type: " & TypeName(key)
		'p_Trace "FindEx(" & Key & ") starting: lo=" & lo & ", hi=" & hi
		While hi - lo > 0
			Dim m
			m = lo + (hi - lo) \ 2
			'p_Trace "FindEx(" & Key & ") comparing with Index " & m & " (" & m_keys(m) & "), hi=" & hi & ", lo=" & lo
			If key <= m_keys(m) Then
				'p_Trace "'" & Key & "'<='" & m_keys(m) & "', hi (" & hi & ") -> " & m
				hi = m
			Else
				'p_Trace "'" & Key & "'>'" & m_keys(m) & "',lo (" & lo & ") -> " & m + 1
				lo = m + 1
			End If
		Wend
		'p_Trace "FindEx(" & Key & ") done lo=" & lo & ", hi=" & hi '& ", ub=" & UB
		If bExact Then
			If lo >= Count Then
				lo = -1 ' not found
			ElseIf m_keys(lo) <> Key Then
				lo = -1 ' not found
			End If
		End If
		FindEx = lo
		'p_Trace "FindEx(" & Key & ") result=" & FindEx
	End Function

	Public Function Find(ByVal Key) ' default Find() is an exact search
		Find = FindEx(Key, True)
	End Function

	Public Function Exists(ByVal Key)
		Exists = Find(Key) <> -1
	End Function

	Public Sub Add(ByVal Key, ByVal Value)
		Dim Index, tmp
		If TypeName(Key)="Null" Then Key = ""
		If IsObject(Key) Then
			Err.Raise -1, "Dictionary{" & m_debug & "}.Add([" & TypeName(Key) & "])", "Keys cannot be objects ( Value Type=" & TypeName(Value) & ")"
			Exit Sub
		End If
		'If IsNumeric(Key) Then
		'	Err.Raise -1, "Dictionary{" & m_debug & "}.Add(" & Key & ")", "Keys cannot be numeric ( Value Type=" & TypeName(Value) & ")"
		'	Exit Sub
		'End If
		Index = FindEx(Key, False) ' get insertion index
		'p_Trace "Add(" & key & ") got insertion index of " & Index
		If Index < Count And m_keys(Index) = Key Then
			Err.Raise -1, "Dictionary{" & m_debug & "}.Add(" & Key & ")", "Key already exists"
			Exit Sub
		End If
		tmp = Count
		Count = tmp + 1

		Dim i
		For i = Count-1 To Index Step -1
			m_keys(i+1) = m_keys(i)
			p_Assign m_data(i+1), m_data(i)
		Next

		m_keys(Index) = Key
		p_Assign m_data(Index), Value
		'p_Trace "Add(" & key & ") DONE! UBound(m_keys)=" & UBound(m_keys) & ", UBound(m_data)=" & UBound(m_data) & ", Count=" & Count
	End Sub

	Public Function Remove(ByVal IndexKey)
		Dim Index, i, tmp
		If IsNumeric(IndexKey) Then
			Index = CLng(IndexKey)
		Else
			Index = Find(IndexKey)
		End If
		If -1 = Index Then
			Remove = Empty
			Err.Raise -1, "Dictionary{" & m_debug & "}.Remove(" & IndexKey & ")", "Invalid Index/Key:  """ & IndexKey & """"
			Exit Function
		End If
		p_Assign Remove, m_data(Index)

		For i = Index To Count-2
			m_keys(i) = m_keys(i+1)
			p_Assign m_data(i), m_data(i+1)
		Next
		m_keys(i) = Empty
		m_data(i) = Empty

		Count = Count - 1
	End Function

	Public Property Get Keys()
		If Count = 0 Then
			Keys = Array()
			Exit Property
		End If
		Dim ar
		ar = m_keys
		ReDim Preserve ar(Count - 1)
		Keys = ar
	End Property

	Public Property Get Values()
		If Count = 0 Then
			Values = Array()
			Exit Property
		End If
		Dim ar
		ar = m_data
		ReDim Preserve ar(Count - 1)
		Values = ar
	End Property
	
	Public Function RemoveAll()
		RemoveAll = Clear()
	End Function
	
	Public Function Clear()
		Clear = Count
		Count = 0
		m_keys(0) = Empty
		m_data(0) = Empty
	End Function

	Public Property Get DebugName()
	    DebugName = m_debug
	End Property

	Public Property Let DebugName(ByVal value)
		'Err.Raise -1, "Dictionary.DebugName.Let()", "DebugName feature is disabled"
		m_debug = Trim(value & "")
	End Property

	' private function(s) and data

	Private Sub p_Trace(ByVal msg)
		If m_debug <> "" Then Response.Write "<font color=red>" & Now & " " & Server.HtmlEncode(m_debug) & ".Dictionary.'p_Trace(): " & Server.HtmlEncode(msg) & "</font><br>"
	End Sub

	Private Sub p_Assign(ByRef dst, ByVal Value)
		If IsObject(Value) Then
			Set dst = Value
		Else
			dst = Value
		End If
	End Sub

	Private Sub Class_Initialize()
		m_debug = "(unnamed)"
		m_count = 0
		ReDim m_keys(0)
		ReDim m_data(0)
		m_default = Empty
	End Sub

	Private Sub Class_Terminate()
		Erase m_keys
		Erase m_data
	End Sub


End Class
%>