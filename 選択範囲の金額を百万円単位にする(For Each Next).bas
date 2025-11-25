Sub 選択範囲の金額を百万円単位にする()
    
    Dim myRange As Range

    For Each myRange In Selection
        If myRange.Value <> 0 And myRange.Value <> "" And _
        TypeName(myRange.Value) <> "String" And TypeName(myRange.Value) <> "Date" Then
        'セルの値が0,空白,文字列,日付のいづれかの場合は処理をしない
            myRange.Value = Application.RoundUp(myRange.Value / 1000000, 2)    '小数点第三位以下切り上げ
        End If
    Next myRange
    
End Sub