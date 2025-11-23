Sub FL列にチェックを記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = ChrW(10003)

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            Cells(myRange.Row, "FL").Value = V
            
        End If
    Next myRange

End Sub
Sub FM列にチェックを記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = ChrW(10003)

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            Cells(myRange.Row, "FM").Value = V
            
        End If
    Next myRange

End Sub