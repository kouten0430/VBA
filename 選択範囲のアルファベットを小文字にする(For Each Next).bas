Sub 選択範囲のアルファベットを小文字にする()
    
    Dim myRange As Range

    For Each myRange In Selection
        myRange.Value = StrConv(myRange.Value, vbLowerCase)
    Next myRange
    
End Sub