Sub 選択範囲の金額を千円単位にする四捨五入版()
    
    Dim myRange As Range

    For Each myRange In Selection
        If myRange.Value <> 0 And myRange.Value <> "" And _
        TypeName(myRange.Value) <> "String" And TypeName(myRange.Value) <> "Date" Then
        'セルの値が0,空白,文字列,日付のいづれかの場合は処理をしない
            myRange.Value = Round(myRange.Value / 1000, 0)    '小数点以下は四捨五入
        End If
    Next myRange
    
End Sub