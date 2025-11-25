Sub 冒頭に建災防を追加する()
    'ローカル的なマクロです
    Dim V As Variant
    Dim myRange As Range
    
    V = "建災防　"
    
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V & myRange.Value
        End If
    Next myRange

End Sub
Sub 冒頭に研修センターを追加する()
    'ローカル的なマクロです
    Dim V As Variant
    Dim myRange As Range
    
    V = "研修センター　"
    
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V & myRange.Value
        End If
    Next myRange

End Sub
Sub 冒頭に401会議室を追加する()
    'ローカル的なマクロです
    Dim V As Variant
    Dim myRange As Range
    
    V = "401会議室　"
    
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V & myRange.Value
        End If
    Next myRange

End Sub
Sub 冒頭に402会議室を追加する()
    'ローカル的なマクロです
    Dim V As Variant
    Dim myRange As Range
    
    V = "402会議室　"
    
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V & myRange.Value
        End If
    Next myRange

End Sub
Sub 冒頭に701会議室を追加する()
    'ローカル的なマクロです
    Dim V As Variant
    Dim myRange As Range
    
    V = "701会議室　"
    
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V & myRange.Value
        End If
    Next myRange

End Sub
Sub 冒頭に小会議室を追加する()
    'ローカル的なマクロです
    Dim V As Variant
    Dim myRange As Range
    
    V = "小会議室　"
    
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V & myRange.Value
        End If
    Next myRange

End Sub
Sub 冒頭に安全衛生教育センターを追加する()
    'ローカル的なマクロです
    Dim V As Variant
    Dim myRange As Range
    
    V = "安全衛生教育センター　"
    
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V & myRange.Value
        End If
    Next myRange

End Sub
Sub 冒頭に本店を追加する()
    'ローカル的なマクロです
    Dim V As Variant
    Dim myRange As Range
    
    V = "本店　"
    
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V & myRange.Value
        End If
    Next myRange

End Sub
Sub 冒頭にTV会議を追加する()
    'ローカル的なマクロです
    Dim V As Variant
    Dim myRange As Range
    
    V = "TV会議　"
    
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V & myRange.Value
        End If
    Next myRange

End Sub