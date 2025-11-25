Sub 創を記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "創"

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V
            
        End If
    Next myRange

End Sub
Sub カを記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "カ"

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V
            
        End If
    Next myRange

End Sub
Sub 特を記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "特"

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V
            
        End If
    Next myRange

End Sub
Sub ゆを記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "ゆ"

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V
            
        End If
    Next myRange

End Sub
Sub リを記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "リ"

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V
            
        End If
    Next myRange

End Sub
Sub A半を記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "A半"

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V
            
        End If
    Next myRange

End Sub
Sub P半を記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "P半"

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V
            
        End If
    Next myRange

End Sub