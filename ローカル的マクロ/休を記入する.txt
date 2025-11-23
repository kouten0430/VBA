Sub 休を記入する()
'記入後に文字色を赤にします
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "休"

    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V
            
        End If
    Next myRange
    
    Selection.Font.Color = 255

End Sub