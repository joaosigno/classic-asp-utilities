<%
' Any copyright is dedicated to the Public Domain.
' http://creativecommons.org/publicdomain/zero/1.0/
Const SBS_None = 0 ' last one

' SELECT states
Const SBS_Sel_Select = 10
Const SBS_Sel_From = 11
Const SBS_Sel_InnerJoin = 12
Const SBS_Sel_On = 13
Const SBS_Sel_Where = 14
Const SBS_Sel_GroupBy = 15
Const SBS_Sel_Having = 16
Const SBS_Sel_OrderBy = 17

Const SBS_Ins_InsertInto = 20
Const SBS_Ins_Values = 21

Const SBS_Upd_Update = 30
Const SBS_Upd_Where = 31

Const SBS_Del_DeleteFrom = 40
Const SBS_Del_Where = 41

' TODO FIXME: Create and Drop support


Class SqlBuilder
	Private mmE, mmState, mmVals
	Private mmTable, mmTable2, mmGlue
	Public mSql, mSql2, mSql3

	Public Sub mInit ( ByRef pSqlEncoder )
		Set mmE = pSqlEncoder
		Set mmVals = CreateObject("Scripting.Dictionary")
		mmNewState SBS_None, Empty
		mSql = ""
		mSql2 = ""
		mSql3 = ""
		mmGlue = ""
	End Sub

	Private Function mmNewState ( ByVal pNewState, ByVal pValidOldStates )
		If Not IsEmpty(pValidOldStates) Then
			mmAssertState pValidOldStates
		End If
		mmState = pNewState
		mmNewState = mmState
	End Function

	Public Sub mField ( ByVal pName )
		mmAddField pName, ""
	End Sub
	Public Sub mFieldAs ( ByVal pExpr, ByVal pAs )
		mmAddField pAs, pExpr
	End Sub
	Public Sub mString ( ByVal pName, ByVal pValue )
		mmAddVal pName, mmE.String(pValue)
	End Sub
	Public Sub mNumeric ( ByVal pName, ByVal pValue )
		mmAddVal pName, mmE.Numeric(pValue)
	End Sub
	Public Sub mDateOnly ( ByVal pName, ByVal pValue )
		mmAddVal pName, mmE.DateOnly(pValue)
	End Sub
	Public Sub mDateTime ( ByVal pName, ByVal pValue )
		mmAddVal pName, mmE.DateTime(pValue)
	End Sub
	Public Sub mGuess ( ByVal pName, ByVal pValue )
		mmAddVal pName, mmE.Guess(pValue)
	End Sub
	Public Sub mField_String ( ByVal pName, ByVal pValue )
		mField pName
		mString pName, pValue
	End Sub
	Public Sub mField_Numeric ( ByVal pName, ByVal pValue )
		mField pName
		mNumeric pName, pValue
	End Sub
	Public Sub mField_DateOnly ( ByVal pName, ByVal pValue )
		mField pName
		mDateOnly pName, pValue
	End Sub
	Public Sub mField_DateTime ( ByVal pName, ByVal pValue )
		mField pName
		mDateTime pName, pValue
	End Sub
	Public Sub mField_Guess ( ByVal pName, ByVal pValue )
		mField pName
		mGuess pName, pValue
	End Sub
	Public Sub mFieldAs_String ( ByVal pExpr, ByVal pAs, ByVal pValue )
		mFieldAs pExpr, pAs
		mString pAs, pValue
	End Sub
	Public Sub mFieldAs_Numeric ( ByVal pExpr, ByVal pAs, ByVal pValue )
		mFieldAs pExpr, pAs
		mNumeric pAs, pValue
	End Sub
	Public Sub mFieldAs_DateOnly ( ByVal pExpr, ByVal pAs, ByVal pValue )
		mFieldAs pExpr, pAs
		mDateOnly pAs, pValue
	End Sub
	Public Sub mFieldAs_DateTime ( ByVal pExpr, ByVal pAs, ByVal pValue )
		mFieldAs pExpr, pAs
		mDateTime pAs, pValue
	End Sub
	Public Sub mFieldAs_Guess ( ByVal pExpr, ByVal pAs, ByVal pValue )
		mFieldAs pExpr, pAs
		mGuess pAs, pValue
	End Sub

	Private Sub mmAddField ( ByVal pField, ByVal pExpr )
		Select Case mmState
		Case SBS_Sel_Select
			mSql = mSql & mmGlue
			If pExpr <> "" Then
				mSql = mSql & " " & pExpr & " as"
			End If
			mSql = mSql & " [" & pField & "]"
			mmGlue = ","
		Case SBS_Ins_InsertInto
			If Left(mSql2,7) = " values" Then
				mSql2 = " )" & mSql2
				mSql = mSql & " ("
			Else
				mSql = mSql & ","
			End If
			mSql = mSql & " [" & pField & "]"
		Case SBS_Upd_Update
			mSql = mSql & mmGlue & " [" & pField & "]"
		End Select
	End Sub
	Private Sub mmAddVal ( ByVal pKey, ByVal pValue )
		mmVals ( LCase(Trim(pKey)) ) = pValue
		Select Case mmState
		Case SBS_Ins_InsertInto, SBS_Ins_Values
			mSql2 = mSql2 & mmGlue & " {" & pKey & "}"
			mmGlue = ","
		Case SBS_Upd_Update
			mSql = mSql & "={" & pKey & "}"
			mmGlue = ","
		End Select
	End Sub

	Public Sub mSelect ( ByVal pList )
		mmAssertState SBS_None
		mmState = SBS_Sel_Select
		mSql = "select"
		If pList <> "" Then
			mSql = mSql & " " & pList
			mmGlue = ","
		End If
	End Sub

	Public Sub mFrom ( ByVal pTable )
		mmNewState SBS_Sel_From, SBS_Sel_Select
		mmTable = pTable
		'mmAddField pTable, ""
		mSql = mSql & " from [" & pTable & "]"
	End Sub

	Public Sub mInnerJoin ( ByVal pTable )
		mmNewState SBS_Sel_InnerJoin, SBS_Sel_From
		mmTable2 = pTable
		'mmAddField pTable, ""
		mSql = mSql & " inner join [" & pTable & "]"
	End Sub

	Public Sub mOn ( ByVal pKey1, ByVal pKey2 )
		If Left(pKey1,1) <> "[" Then pKey1 = "[" & pKey1 & "]"
		If Left(pKey2,1) <> "[" Then pKey2 = "[" & pKey2 & "]"
		Select Case mmState
		Case SBS_Sel_InnerJoin
			mSql = mSql & " on"
		Case SBS_Sel_On
			mSql = mSql & ","
		End Select
		mmNewState SBS_Sel_On, Array ( SBS_Sel_InnerJoin, SBS_Sel_On )
		mSql = mSql & " [" & mmTable & "]." & pKey1 & "=[" & mmTable2 & "]." & pKey2
	End Sub

	Public Sub mWhere ( ByVal pCond )
		Select Case mmAssertState ( Array ( SBS_Sel_From, SBS_Sel_On, SBS_Upd_Update, SBS_Del_DeleteFrom ) )
		Case SBS_Sel_From, SBS_Sel_On
			mmState = SBS_Sel_Where
		Case SBS_Upd_Update
			mmState = SBS_Upd_Where
		Case SBS_Del_DeleteFrom
			mmState = SBS_Del_Where
		End Select
		If pCond = "" Then Err.Raise -1, TypeName(Me) & ".mWhere", "parameter `pCond` is required"
		mSql = mSql & " where " & pCond
	End Sub

	Public Sub mGroupBy ( ByVal pFields )
		mmNewState SBS_Sel_GroupBy, Array ( SBS_Sel_From, SBS_Sel_On, SBS_Sel_Where )
		If IsArray(pFields) Then pFields = Join(pFields,",")
		If pFields = "" Then Err.Raise -1, TypeName(Me) & ".mGroupBy", "parameter `pFields` is required"
		' if user passed [] around fields, get rid of them so we can reinject them and ensure quality...
		pFields = Replace(Replace(Replace(Replace(pFields,"[",""),"]",""),".","].["),",","],[")
		mSql = mSql & " group by [" & pFields & "]"
	End Sub

	Public Sub mHaving ( ByVal pCond )
		mmNewState SBS_Sel_Having, SBS_Sel_GroupBy
		If pCond = "" Then Err.Raise -1, TypeName(Me) & ".mHaving", "parameter `pCond` is required"
		mSql = mSql & " having " & pCond
	End Sub

	Public Sub mOrderBy ( ByVal pFields )
		mmNewState SBS_Sel_OrderBy, Array ( SBS_Sel_From, SBS_Sel_On, SBS_Sel_Where, SBS_Sel_GroupBy, SBS_Sel_Having )
		If IsArray(pFields) Then pFields = Join(pFields,",")
		If pFields = "" Then Err.Raise -1, TypeName(Me) & ".mOrderBy", "parameter `pFields` is required"
		mSql = mSql & " order by [" & Replace(pFields,",","],[") & "]"
	End Sub

	Public Sub mInsertInto ( ByVal pTable )
		mmNewState SBS_Ins_InsertInto, SBS_None
		mmTable = pTable
		'mmAddField pTable, ""
		mSql = "insert into [" & pTable & "]"
		mSql2 = " values ("
		mSql3 = " )"
		mmGlue = ""
	End Sub

	Public Sub mUpdate ( ByVal pTable )
		mmNewState SBS_Upd_Update, SBS_None
		mmTable = pTable
		'mmAddField pTable, ""
		mSql = "update [" & pTable & "] set"
		mmGlue = ""
	End Sub

	Public Sub mValues()
		mmNewState SBS_Ins_Values, SBS_Ins_InsertInto
	End Sub

	Public Sub mDeleteFrom ( ByVal pTable )
		mmNewState SBS_Del_DeleteFrom, SBS_None
		mmTable = pTable
		'mmAddField pTable, ""
		mSql = "delete from [" & pTable & "]"
	End Sub

	Public Function mBuild()
		Err.Source = TypeName(Me) & ".Build"
		Dim sql, out, i, last, n, match, tmp, j, ar
		sql = mSql & mSql2 & mSql3
		last = Len(sql)
		i = 1
		While i <= last
			match = 1
			tmp = Mid(sql,i,1)
			Select Case tmp
			Case "'" ' string
				n = InStr(i+1,sql,"'")
				If n > 0 Then
					match = n - i + 1
					tmp = Mid(sql,i,match)
				End If
			Case "[" ' field name
				n = InStr(i+1,sql,"]")
				If n > 0 Then
					match = n - i + 1
					ar = Split(Mid(sql,i+1,match-2),".")
					tmp = ""
					For j = 0 To UBound(ar)
						If j > 0 Then tmp = tmp & "."
						tmp = tmp & mmE.Field(ar(j))
					Next
				End If
			Case "{"
				n = InStr(i+1,sql,"}")
				If n > 0 Then
					match = n - i + 1
					tmp = LCase(Trim(Mid(sql,i+1,match-2)))
					If Not mmVals.Exists(tmp) Then
						Err.Raise -1,, "named value doesn't exist: " & tmp '& ", the following named values do exist: " & Join(mmVals.Keys,",")
					Else
						tmp = mmVals ( tmp )
					End If
				End If
			'Case Else
				' TODO: search for next special-case...
			End Select
			out = out & tmp
			i = i + match
		Wend
		mBuild = out
	End Function

	Private Function mmAssertState ( ByVal expected )
		mmAssertState = mmState
		If Not IsArray(expected) Then
			If mmState = expected Then Exit Function
			Err.Raise -1, TypeName(Me) & ".mmAssertState", "Invalid attempted state transition from " & mmState & " to " & expected
		Else
			Dim i
			For i = 0 To UBound(expected)
				If mmState = expected(i) Then Exit Function
			Next
			Err.Raise -1, TypeName(Me) & ".mmAssertState", "Invalid attempted state transition from " & mmState & " to " & Join(expected," or ")
		End If
	End Function
End Class

Class SqlInsertBuilder
	Private conf, e
	Private TableName
	Private Flds, Vals
	
	Public Sub Init ( ByRef pConf, ByVal pTableName )
		Set conf = pConf
		Set e = conf.SqlEncoder
		TableName = pTableName
	End Sub
	Public Sub AddNumeric ( ByVal pName, ByVal pValue )
		mAdd pName, e.Numeric(pValue)
	End Sub
	Public Sub AddString ( ByVal pName, ByVal pValue )
		mAdd pName, e.String(pValue)
	End Sub
	Public Sub AddDateOnly ( ByVal pName, ByVal pValue )
		mAdd pName, e.DateOnly(pValue)
	End Sub
	Public Sub AddDateTime ( ByVal pName, ByVal pValue )
		mAdd pName, e.DateTime(pValue)
	End Sub
	Public Sub AddGuess ( ByVal pName, ByVal pValue )
		mAdd pName, e.Guess(pValue)
	End Sub
	Private Sub mAdd ( ByVal pName, ByVal pValue )
		Flds = Flds & "," & e.Field(pName)
		Vals = Vals & "," & pValue
	End Sub
	Public Function Execute()
		Set Execute = conf.Execute ( "insert into {0_} ( " & Mid(Flds,2) & " ) values ( " & Mid(Vals,2) & " );", _
			TableName )
	End Function
End Class

Class SqlInnerJoinBuilder
	Private mConf, mLeft, mRight, mFlds, mVals
	Public sql
	
	Public Sub Init ( ByRef pConf )
		Set mConf = pConf
		Set mFlds = CreateObject("Scripting.Dictionary")
		Set mVals = CreateObject("Scripting.Dictionary")
		sql = ""
	End Sub
	
	Public Function Left ( ByVal pTable )
		Set mLeft = New SqlInnerJoinBuilderTable
		mLeft.Init Me, pTable
		Set Left = mLeft
	End Function

	Public Function Right ( ByVal pTable )
		Set mRight = New SqlInnerJoinBuilderTable
		mRight.Init Me, pTable
		Set Right = mRight
	End Function
	
	Public Sub Field ( ByVal pTable, ByVal pName )
		mFlds(pName) = pTable
		sql = sql & "," & conf.SqlFormat ( "{0_}.{1_}", Array ( pTable, pName ) )
	End Sub
	
	Public Sub Join ( ByVal pLeftKey, ByVal pRightKey )
		sql = "select " & Mid(sql,2) & conf.SqlFormat(" from {0_} inner join {1_} on {0_}.{2_}={1_}.{3_}", _
			Array ( mLeft.Table, mRight.Table, pLeftKey, pRightKey ) )
	End Sub
	
	Public Sub Value ( ByVal pName, ByVal pValue )
		mVals(pName) = pValue
	End Sub
	
	Public Sub Finish ( Byval append )
		Dim Key
		For Each Key In mFlds
			append = Replace(append,"[" & Key & "]",conf.SqlFormat("{0_}.{1_}",Array(mFlds(Key),Key)))
		Next
		For Each Key In mVals
			append = Replace(append,"{" & Key & "}",conf.SqlFormat("{0?}",mVals(Key)))
		Next
		sql = sql & " " & append
	End Sub
End Class

' THE FOLLOWING IS A HELPER CLASS, DO NOT USE IT DIRECTLY
Class SqlInnerJoinBuilderTable
	Private mB
	Public Table

	Public Sub Init ( ByRef pB, ByVal pTable )
		Set mB = pB
		Table = pTable
	End Sub
	
	Public Sub Field ( ByVal pName )
		mB.Field Table, pName
	End Sub
End Class

%>
