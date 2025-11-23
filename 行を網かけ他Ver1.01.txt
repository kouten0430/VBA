Sub いいい行を網かけ()
    '行のデータがある範囲を網かけします
    Dim 始 As Long
    Dim 終 As Long
    Dim myRange As Range
    
    If Selection.Address = Selection.EntireRow.Address Then
        Selection.Interior.Pattern = xlPatternGray8
        
    ElseIf Selection.Address = Selection.EntireColumn.Address Then
        MsgBox "列全体が選択されています"
        Exit Sub

    Else
        If Cells(Selection.Row, 1).Value <> "" Then
            始 = 1
        Else
            始 = Cells(Selection.Row, 1).End(xlToRight).Column
        End If
        
        If Cells(Selection.Row, Columns.Count).Value <> "" Then
            終 = Columns.Count
        Else
            終 = Cells(Selection.Row, Columns.Count).End(xlToLeft).Column
        End If
        
        For Each myRange In Selection
            Range(Cells(myRange.Row, 始), Cells(myRange.Row, 終)).Interior.Pattern = xlPatternGray8
            
        Next myRange
        
    End If

End Sub
Sub いいい行の網かけをクリア()
    Dim 始 As Long
    Dim 終 As Long
    Dim myRange As Range
    
    If Selection.Address = Selection.EntireRow.Address Then
        Selection.Interior.Pattern = xlPatternSolid
        
    ElseIf Selection.Address = Selection.EntireColumn.Address Then
        MsgBox "列全体が選択されています"
        Exit Sub

    Else
        If Cells(Selection.Row, 1).Value <> "" Then
            始 = 1
        Else
            始 = Cells(Selection.Row, 1).End(xlToRight).Column
        End If
        
        If Cells(Selection.Row, Columns.Count).Value <> "" Then
            終 = Columns.Count
        Else
            終 = Cells(Selection.Row, Columns.Count).End(xlToLeft).Column
        End If
        
        For Each myRange In Selection
            Range(Cells(myRange.Row, 始), Cells(myRange.Row, 終)).Interior.Pattern = xlPatternSolid
            
        Next myRange
        
    End If
    
End Sub
Sub いいい列を網かけ()
    '列のデータがある範囲を網かけします
    Dim 始 As Long
    Dim 終 As Long
    Dim myRange As Range
    
    If Selection.Address = Selection.EntireColumn.Address Then
        Selection.Interior.Pattern = xlPatternGray8
        
    ElseIf Selection.Address = Selection.EntireRow.Address Then
        MsgBox "行全体が選択されています"
        Exit Sub

    Else
        If Cells(1, Selection.Column).Value <> "" Then
            始 = 1
        Else
            始 = Cells(1, Selection.Column).End(xlDown).Row
        End If
        
        If Cells(Rows.Count, Selection.Column).Value <> "" Then
            終 = Rows.Count
        Else
            終 = Cells(Rows.Count, Selection.Column).End(xlUp).Row
        End If
        
        For Each myRange In Selection
            Range(Cells(始, myRange.Column), Cells(終, myRange.Column)).Interior.Pattern = xlPatternGray8
            
        Next myRange
        
    End If
    
End Sub
Sub いいい列の網かけをクリア()
    Dim 始 As Long
    Dim 終 As Long
    Dim myRange As Range
    
    If Selection.Address = Selection.EntireColumn.Address Then
        Selection.Interior.Pattern = xlPatternSolid
        
    ElseIf Selection.Address = Selection.EntireRow.Address Then
        MsgBox "行全体が選択されています"
        Exit Sub

    Else
        If Cells(1, Selection.Column).Value <> "" Then
            始 = 1
        Else
            始 = Cells(1, Selection.Column).End(xlDown).Row
        End If
        
        If Cells(Rows.Count, Selection.Column).Value <> "" Then
            終 = Rows.Count
        Else
            終 = Cells(Rows.Count, Selection.Column).End(xlUp).Row
        End If
        
        For Each myRange In Selection
            Range(Cells(始, myRange.Column), Cells(終, myRange.Column)).Interior.Pattern = xlPatternSolid
            
        Next myRange
        
    End If

End Sub