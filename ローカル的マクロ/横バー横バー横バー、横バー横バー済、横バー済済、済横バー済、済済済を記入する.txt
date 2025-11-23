Sub 横バー横バー横バーを記入する()
    '記入したい範囲の左端のみ選択して実行する（複数選択可）
    'ローカル的なマクロです
    Dim 配列 As Variant
    Dim myRange As Range
    
    配列 = Array("－", "－", "－")
    
    For Each myRange In Selection
        myRange.Resize(1, UBound(配列) + 1).Value = 配列
        
    Next myRange

End Sub
Sub 横バー横バー済を記入する()
    '記入したい範囲の左端のみ選択して実行する（複数選択可）
    'ローカル的なマクロです
    Dim 配列 As Variant
    Dim myRange As Range
    
    配列 = Array("－", "－", "済")
    
    For Each myRange In Selection
        myRange.Resize(1, UBound(配列) + 1).Value = 配列
        
    Next myRange

End Sub
Sub 横バー済済を記入する()
    '記入したい範囲の左端のみ選択して実行する（複数選択可）
    'ローカル的なマクロです
    Dim 配列 As Variant
    Dim myRange As Range
    
    配列 = Array("－", "済", "済")
    
    For Each myRange In Selection
        myRange.Resize(1, UBound(配列) + 1).Value = 配列
        
    Next myRange

End Sub
Sub 済横バー済を記入する()
    '記入したい範囲の左端のみ選択して実行する（複数選択可）
    'ローカル的なマクロです
    Dim 配列 As Variant
    Dim myRange As Range
    
    配列 = Array("済", "－", "済")
    
    For Each myRange In Selection
        myRange.Resize(1, UBound(配列) + 1).Value = 配列
        
    Next myRange

End Sub
Sub 済済済を記入する()
    '記入したい範囲の左端のみ選択して実行する（複数選択可）
    'ローカル的なマクロです
    Dim 配列 As Variant
    Dim myRange As Range
    
    配列 = Array("済", "済", "済")
    
    For Each myRange In Selection
        myRange.Resize(1, UBound(配列) + 1).Value = 配列
        
    Next myRange

End Sub