Sub I列に横バー横バー横バーを記入する()
    'ローカル的なマクロです
    Dim 配列 As Variant
    Dim myRange As Range
    
    配列 = Array("－", "－", "－")
    
    For Each myRange In Selection
        Cells(myRange.Row, "I").Resize(1, UBound(配列) + 1).Value = 配列
        
    Next myRange

End Sub
Sub I列に横バー横バー済を記入する()
    'ローカル的なマクロです
    Dim 配列 As Variant
    Dim myRange As Range
    
    配列 = Array("－", "－", "済")
    
    For Each myRange In Selection
        Cells(myRange.Row, "I").Resize(1, UBound(配列) + 1).Value = 配列
        
    Next myRange

End Sub
Sub I列に横バー済済を記入する()
    'ローカル的なマクロです
    Dim 配列 As Variant
    Dim myRange As Range
    
    配列 = Array("－", "済", "済")
    
    For Each myRange In Selection
        Cells(myRange.Row, "I").Resize(1, UBound(配列) + 1).Value = 配列
        
    Next myRange

End Sub
Sub I列に済横バー済を記入する()
    'ローカル的なマクロです
    Dim 配列 As Variant
    Dim myRange As Range
    
    配列 = Array("済", "－", "済")
    
    For Each myRange In Selection
        Cells(myRange.Row, "I").Resize(1, UBound(配列) + 1).Value = 配列
        
    Next myRange

End Sub
Sub I列に済済済を記入する()
    'ローカル的なマクロです
    Dim 配列 As Variant
    Dim myRange As Range
    
    配列 = Array("済", "済", "済")
    
    For Each myRange In Selection
        Cells(myRange.Row, "I").Resize(1, UBound(配列) + 1).Value = 配列
        
    Next myRange

End Sub