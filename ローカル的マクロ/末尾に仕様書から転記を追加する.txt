Sub 末尾に仕様書から転記を追加する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "※日時は仕様書から転記"

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = myRange.Value & V
            
        End If
    Next myRange

End Sub