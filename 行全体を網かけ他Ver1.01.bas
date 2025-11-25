Sub 行全体を網かけ()
    Dim myRange As Range
    
    If Selection.Address = Selection.EntireRow.Address Then
        Selection.Interior.Pattern = xlPatternGray8
        
    ElseIf Selection.Address = Selection.EntireColumn.Address Then
        MsgBox "列全体が選択されています"
        Exit Sub

    Else
        For Each myRange In Selection
            myRange.EntireRow.Interior.Pattern = xlPatternGray8
            
        Next myRange
        
    End If

End Sub
Sub 行全体の網かけをクリア()
    Dim myRange As Range
    
    If Selection.Address = Selection.EntireRow.Address Then
        Selection.Interior.Pattern = xlPatternSolid
        
    ElseIf Selection.Address = Selection.EntireColumn.Address Then
        MsgBox "列全体が選択されています"
        Exit Sub

    Else
        For Each myRange In Selection
            myRange.EntireRow.Interior.Pattern = xlPatternSolid
            
        Next myRange
        
    End If
    
End Sub
Sub 列全体を網かけ()
    Dim myRange As Range
    
    If Selection.Address = Selection.EntireColumn.Address Then
        Selection.Interior.Pattern = xlPatternGray8
        
    ElseIf Selection.Address = Selection.EntireRow.Address Then
        MsgBox "行全体が選択されています"
        Exit Sub

    Else
        For Each myRange In Selection
            myRange.EntireColumn.Interior.Pattern = xlPatternGray8
            
        Next myRange
        
    End If
    
End Sub
Sub 列全体の網かけをクリア()
    Dim myRange As Range
    
    If Selection.Address = Selection.EntireColumn.Address Then
        Selection.Interior.Pattern = xlPatternSolid
        
    ElseIf Selection.Address = Selection.EntireRow.Address Then
        MsgBox "行全体が選択されています"
        Exit Sub

    Else
        For Each myRange In Selection
            myRange.EntireColumn.Interior.Pattern = xlPatternSolid
            
        Next myRange
        
    End If

End Sub