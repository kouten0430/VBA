Sub 選択範囲のアルファベットを大文字にする()
    
    Dim myRange As Range

    For Each myRange In Selection
        myRange.Value = StrConv(myRange.Value, vbUpperCase)
    Next myRange
    
End Sub