Sub 選択範囲の文字を全角にする()
    
    Dim myRange As Range

    For Each myRange In Selection
        If IsNumeric(myRange.Value) Then    'セルの値が数字であれば強制的に文字列にする
            myRange.Value = "'" & StrConv(myRange.Value, vbWide)
        Else
            myRange.Value = StrConv(myRange.Value, vbWide)
        End If
    Next myRange
    
End Sub
