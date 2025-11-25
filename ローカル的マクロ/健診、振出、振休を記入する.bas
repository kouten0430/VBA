Sub 健診を記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "健診"

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V
            
        End If
    Next myRange

End Sub
Sub 振出を記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "振出"

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V
            
        End If
    Next myRange

End Sub
Sub 振休を記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "振休"

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V
            
        End If
    Next myRange

End Sub