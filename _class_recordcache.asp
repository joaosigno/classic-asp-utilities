<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Function NewRecordCache ( ByRef rs )
	Set NewRecordCache = New RecordCache
	NewRecordCache.Load rs
End Function

' http://www.devarticles.com/c/a/ASP/Developing-a-VBScript-Class-for-an-Extremely-Lightweight-Recordset-Alternative/2/
Class RecordCache
	'Primitive variables for various purposes
	Private bBOF    'Indicates if we are at beginning of record array
	Private bEOF    'Boolean to indicate if we are at end of record array
	Private intRecordCount  'The number of records/rows in the array 
	Private intFieldCount  'The count of fields/columns in the array
	Private intAbsolutePosition  'The ordinal of the current record/row

	'Array variables for containing the data
	Private arDataset    'The array containing the records
	Private arFieldNames  'An array containing the names of the fields

	Private Sub Class_Initialize()
		'Set initial values for all
		bBOF = True
		bEOF = True
		intRecordCount = 0
		intFieldCount = 0
		intAbsolutePosition = -1

		arDataset = Array()  'Load these with empty arrays
		arFieldNames = Array()    
	End Sub

	'Method to load the data from the Recordset
	Public Sub Load(ByRef rs)
		'First check that we actually received a Recordset
		If TypeName(rs) <> "Recordset" Then
			Err.Raise vbObjectError + 99999, "RecordCache:Load()", "Load method requires a Recordset object."
		End If
		On Error Resume Next
		arDataset = rs.GetRows()    'Harvest the Recordset's data
		'Make sure some data was contained actually in the
		'Recordset.  If not, the .GetRows() method above returns
		'an error.  If so, we know there was no data, and set
		'our variables to suit.
 
		If Err.Number <> 0 Then 
			'.GetRows failed
			intRecordCount = -1
			intFieldCount = 0
			bBOF = True
			bEOF = True
			intAbsolutePosition = -1
			'bReadyForUse = False
			'bRsLoaded = False
			Exit Sub 
		End If
		On Error GoTo 0

		'At this point we know we have some data
		bBOF = False
		bEOF = False
		intAbsolutePosition = 0
		'bRsLoaded = True

		intRecordCount = UBound(arDataset, 2)+1  'Get the # of records
		intFieldCount = rs.Fields.Count    'Get the # of fields
		Dim i, arTmpFields()
		ReDim arTmpFields(intFieldCount-1)

		For i = 0 to intFieldCount-1
			arTmpFields(i) = rs.Fields(i).Name
			Execute "fld_" & rs.fields(i).Name & " = " & i
		Next
		arFieldNames = arTmpFields
		'bReadyForUse = True
	End Sub
	
	Private Sub p_SetAbs ( ByVal value, ByVal bAllowEOF )
		If RecordCount <= 0 Then
			Err.Raise vbObjectError + 99999, "RecordCache:Move" , "RecordCache contains no data"
		End If

		If Not IsNumeric(value) Then
			Err.Raise vbObjectError + 99999, "RecordCache:Move" , "Move method requires an integer"
		End If

		value = CLng(value)
		
		If value > RecordCount Then
			If bAllowEOF Then
				value = RecordCount+1
			Else
				Err.Raise -1, "RecordCache.p_SetAbs()", "Invalid parameter: requested record too high"
				Exit Sub
			End If
		ElseIf value <= 0 Then
			If bAllowEOF Then
				value = 0
			Else
				Err.Raise -1, "RecordCache.p_SetAbs()", "Invalid parameter: requested record too low"
				Exit Sub
			End If
		End If
		intAbsolutePosition = value
	End Sub

	Public Property Get AbsolutePosition()
		AbsolutePosition = intAbsolutePosition
	End Property
	
	Public Property Let AbsolutePosition ( ByVal value )
		p_SetAbs value, False
	End Property
	
	' With ADODB.Recordset.Move(), NumRecords is relative to current position if the 2nd parameter
	' is unspecified. There's not great need or call for 2nd parameter, so we're going to assume
	' always relative
	Public Sub Move(ByVal NumRecords)
		p_SetAbs AbsolutePosition + NumRecords, True
	End Sub
	
	'Move from the current record to the next one.
	Public Sub MoveNext()
		Move 1
	End Sub
	
	'Move from the current record to the previous one.
	Public Sub MovePrevious()
		Move -1
	End Sub

	'Move from the current record to the first record.
	Public Sub MoveFirst()
		AbsolutePosition = 1
	End Sub

	'Move from the current record to the last record.
	Public Sub MoveLast()
		AbsolutePosition = RecordCount
	End Sub
	
	Public Property Get RecordCount()
		RecordCount = intRecordCount
	End Property
	
	'Public Property Get CursorLocation()
	'	CursorLocation = intCursorLocation
	'End Property
	
	'Public Property Let CursorLocation ( ByVal intRecordNumber )
	'	Move ( intRecordNumber )
	'End Property

	Public Property Get EOF()
		EOF = (RecordCount <= 0) Or (intAbsolutePosition > RecordCount)
	End Property

	Public Property Get BOF()
		BOF = (RecordCount <= 0) Or (intAbsolutePosition <= 0)
	End Property

	Public Default Function Fields(ByVal intSelectedField)
		If RecordCount <= 0 Then
			Err.Raise vbObjectError + 99999, "RecordCache:Fields", "RecordCache contains no data"
			Set Fields = Nothing
			Exit Function
		End If

		'Make sure we have allowed for the BOF and EOF conditions
		If BOF() Then
			Err.Raise vbObjectError + 99999, "RecordCache:Fields", "Cursor is at BOF. Use MoveFirst method to access fields."
			Exit Function
		End If

		If EOF() Then
			Err.Raise vbObjectError + 99999, "RecordCache:Fields" , "Cursor is at EOF. Use MoveFirst method to access fields."
			Exit Function
		End If

		Dim intOrdinal

		'Determine if the input is numeric…

		If IsNumeric(intSelectedField) Then
			'It IS numeric, so if its within the 
			'upper and lower ranges, return the data
			intOrdinal = Cint(intSelectedField)
			'Make sure the numeric input is within the range of
			'zero thru the # of fields
			If (intOrdinal > intFieldCount-1) OR (intOrdinal < 0)  Then
				Err.Raise vbObjectError + 99999, "RecordCache:Fields",  "Field #" & intOrdinal & " doesn't exist."
				Exit Function
			End If
		Else
			'It's NOT numeric, so treat it as a string and
			'cycle through the field-name array looking
			'for a match.

			Dim strTest
			strTest = "fld_" & intSelectedField
			If Len( Eval( strTest ) ) = 0 Then
				'No such variable exists, so this is not a valid field name
				Err.Raise vbObjectError + 99999, "RecordCache:Fields", "No field with the name '" & intSelectedField & "' exists."
				Exit Function
			Else
				'Such a variable does exist, so we use Eval to tease out its value
				intOrdinal = Eval( strTest )
			End If
		End If
		'Check to see if any matching ordinal value was found…
		If intOrdinal = -1 Then
			Err.Raise vbObjectError + 99999, "RecordCache:Fields", "No field with the name '" & intSelectedField & "' exists."
			Exit Function
		End If
		'We seem to have a valid field number, so 
		'return the field data 
		Fields = arDataset(intOrdinal, intAbsolutePosition-1)
	End Function
	
	Public Sub Persist(ByVal strMethod, ByVal strUniqueID)
		'To persist these two arrays, we will create
		'another array, with one-dimension and two elements,
		'and load our two existing arrays into it.  Then
		'we will save this new array as needed.
		Dim ar(1), vArray
		ar(0) = arFieldNames
		ar(1) = arDataset
		vArray = ar
		'Persist vArray as a Session or Application variable,
		'creating a key for it with the structure:
		'
		'   RecordCache:[unique-id]

		Select Case strMethod
			Case "SESSION"
				Session("RecordCache:" & strUniqueID) = vArray
			Case "APPLICATION"
				Application("RecordCache:" & strUniqueID) = vArray
			Case Else
				Err.Raise vbObjectError, "RecordCache:Persist", "You must specify either 'Session' or 'Application' as the method."
		End Select
	End Sub

	Public Sub LoadPersisted(ByVal strMethod, ByVal strUniqueID)
		'A new method (incomplete) to load persisted values
		'from file or Session variable to 
		strMethod = Trim( UCase(strMethod) )
		Dim ar
		Select Case strMethod
			Case "SESSION"
				ar = Session("RecordCache:" & strUniqueID)
			Case "APPLICATION"
				' pull data from Application variable
				ar = Application("RecordCache:" & strUniqueID)
			Case Else
				Err.Raise vbObjectError + 99999, "RecordCache:Persist", "You must specify either 'Session' or 'Application' as the method."
		End Select
		If Not IsArray(ar) Then
			Err.Raise vbObjectError + 99999, "RecordCache:Persist", "No persisted RecordCache found with this ID: " & strUniqueID
			Exit Sub
		End If
		'At this point, assume we have successfully acquired a persisted
		'RecordCache variable.  Prepare the other key internal variables.
		arFieldNames = ar(0)
		arDataset = ar(1)        
		bBOF = False
		bEOF = False
		intCursorLocation = 0
		intFieldCount = UBound(arFieldNames)+1
		intRecordCount = UBound( arDataset, 2 )+1
		'bRsLoaded = True
		'bReadyForUse = True
	End Sub

	'Destroys the persisted file or variable.
	Public Sub DePersist(ByVal strMethod, ByVal strUniqueID)
		strMethod = Trim( UCase(strMethod) )
		Select Case strMethod
			Case "SESSION"
				Session("RecordCache:" & strUniqueID) = Empty
			Case "APPLICATION"
				Application("RecordCache:" & strUniqueID) = Empty
			Case Else
				Err.Raise vbObjectError + 99999, "RecordCache:DePersist", "You must specify either 'Session' or 'Application' as the method."
		End Select
	End Sub

	'Determine if such data was persisted
	Public Function IsPersisted(ByVal strMethod, ByVal strUniqueID)
		strMethod = Trim( UCase(strMethod) )

		Select Case strMethod
			Case "SESSION"
				IsPersisted = Not IsEmpty( Session("RecordCache:" & strUniqueID) )
			Case "APPLICATION"
				IsPersisted = Not IsEmpty( Application("RecordCache:" & strUniqueID) )
			Case Else
				Err.Raise vbObjectError + 99999, "RecordCache:IsPersisted ", "You must specify either 'Session' or 'Application' as the method."
		End Select
	End Function

End Class
%>
