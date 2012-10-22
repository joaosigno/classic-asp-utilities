<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Class CsvRecordsetPrivate
	Public data, keys, ar, MaxIndex

	Public Sub Normalize()
		ReDim Preserve ar(MaxIndex)
	End Sub
End Class

Class CsvField
	Public Name, Index, priv

	Public Default Property Get Value()
		Value = priv.ar(Index)
	End Property
End Class

Class CsvFields
	Public priv, Count, arFields

	Public Sub Class_Initialize()
		Count = 0
	End Sub

	Public Default Property Get Items ( ByVal aKeyIndex )
		If Not IsNumeric(aKeyIndex) Then
			aKeyIndex = priv.keys(aKeyIndex)
		End If
		'Response.Write "aKeyIndex=" & aKeyIndex & "<br>"
		'Response.Write "TypeName(arFields(" & aKeyIndex & "))=" & TypeName(arFields(aKeyIndex))
		'Response.Flush
		Set Items = arFields(aKeyIndex)
	End Property
End Class

Class CsvRecordset
	Public priv, lines, Record, RecordCount
	Public Fields

	Public Sub Class_Initialize()
	End Sub

	Public Default Property Get FieldItems ( ByVal aKeyIndex )
		Set FieldItems = Fields.Items(aKeyIndex)
	End Property

	Public Function LoadMemory ( ByRef aData )
		Set priv = New CsvRecordsetPrivate
		priv.data = aData
		Set priv.keys = CreateObject("Scripting.Dictionary")

		Set Fields = New CsvFields
		Fields.Count = 0
		Set Fields.priv = priv

		Dim i_data, len_data, prev, line, i, x

		len_data = Len(priv.data)
		i_data = 1
		prev = 0
		RecordCount = 0
		line = 1
		lines = Array()
		While i_data <= len_data
			priv.ar = CsvSplitEx ( priv.data, i_data )
			If i_data = prev Then
				Err.Raise -1, "", "i_data(" & i_data & ") = prev(" & prev & ")"
			End If
			If 1 = line Then
				priv.MaxIndex = UBound(priv.ar)
				Dim arFields : arFields = Array()
				For i = 0 To priv.MaxIndex
					priv.ar(i) = LCase(Trim(priv.ar(i)))
					If priv.ar(i) <> "" Then
						Set x = New CsvField
						x.Name = priv.ar(i)
						x.Index = i
						Set x.priv = priv
						ReDim Preserve arFields(Fields.Count)
						Set arFields(Fields.Count) = x
						Fields.Count = Fields.Count + 1
					End If
				Next
				Fields.arFields = arFields
			ElseIf Trim(priv.ar(0)) <> "" Or UBound(priv.ar) > 0 Then
				If RecordCount >= UBound(lines) Then
					ReDim Preserve lines(RecordCount+100)
				End If
				lines(RecordCount) = prev
				RecordCount = RecordCount + 1
			End If
			line = line + 1
			prev = i_data
		Wend
		ReDim Preserve lines(RecordCount-1)
		MoveFirst
	End Function

	Public Function LoadFile ( ByVal aFileName )
	End Function

	Public Property Get BOF()
		BOF = (Record < 0)
	End Property

	Public Property Get EOF()
		EOF = (Record >= RecordCount)
	End Property

	Public Sub MoveFirst()
		Record = 0
		Refresh
	End Sub

	Public Sub MoveNext()
		Record = Record + 1
		Refresh
	End Sub

	Public Sub MovePrevious()
		Record = Record - 1
		Refresh
	End Sub

	Public Sub MoveLast()
		Record = RecordCount - 1
		Refresh
	End Sub

	Public Sub Refresh()
		If BOF Or EOF Then
			priv.ar = Array()
		Else
			priv.ar = CsvSplitEx ( priv.data, lines(Record) )
		End If
		priv.Normalize
	End Sub
End Class
%>
