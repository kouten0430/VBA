Sub 二つのセルのデータをZの動きで入れ替えする()
    'ローカル的なマクロです
    Dim i As Integer
    Dim myRange As Range
    Dim 値(1) As String
    Dim 列(1) As Integer
    Dim 行(1) As Long
    
    If Selection.Count <> 2 Or Selection.MergeCells Then Exit Sub

    i = 0
        
    For Each myRange In Selection
        値(i) = myRange.Value
        列(i) = myRange.Column
        行(i) = myRange.Row
        i = i + 1
        
    Next myRange
    
    If 列(0) = 列(1) Or 行(0) = 行(1) Or 値(0) = "" Or 値(1) = "" Then Exit Sub '水平又は垂直に選択している場合、選択位置に空白がある場合は処理しない
    
    i = 1
        
    For Each myRange In Selection
        Cells(行(i), myRange.Column).Value = 値(i)
        i = i - 1
        
    Next myRange
    
    Selection.ClearContents
    
End Sub