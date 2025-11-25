Sub ううう二つのセルのデータをNの動きで入れ替えする待機日直用()
    'ローカル的なマクロです
    Dim i As Integer
    Dim myRange As Range
    Dim 値(1) As String
    Dim 値2(1) As String
    Dim 列(1) As Integer
    Dim 行(1) As Long
    
    If Selection.Count <> 2 Or Selection.MergeCells Then Exit Sub

    i = 0
        
    For Each myRange In Selection
        値(i) = myRange.Value
        値2(i) = myRange.Offset(0, 1).Value
        列(i) = myRange.Column
        行(i) = myRange.Row
        i = i + 1
        
    Next myRange
    
    If 列(0) = 列(1) Or 行(0) = 行(1) Or 値(0) = "" Or 値(1) = "" Then Exit Sub '水平又は垂直に選択している場合、選択位置に空白がある場合は処理しない
    
    Selection.ClearContents
    Selection.Offset(0, 1).ClearContents
    
    i = 1
        
    For Each myRange In Selection
        Cells(myRange.Row, 列(i)).Value = 値(i)
        Cells(myRange.Row, 列(i)).Offset(0, 1).Value = 値2(i)
        i = i - 1
        
    Next myRange
    
End Sub