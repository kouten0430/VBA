Sub 選択範囲のアルファベットの先頭文字を大文字にする()
    
    Dim myRange As Range

    For Each myRange In Selection
        myRange.Value = StrConv(myRange.Value, vbProperCase)
    Next myRange
    
End Sub