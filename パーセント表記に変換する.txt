Sub パーセント表記に変換する()
    '数値をパーセント表記（文字列）に変換する
    Dim myRange As Range

    For Each myRange In Selection
        If myRange.Value <> "" And TypeName(myRange.Value) <> "String" _
        And TypeName(myRange.Value) <> "Date" Then
        'セルの値が空白,文字列,日付のいづれかの場合は処理をしない
            myRange.Value = "'" & myRange.Value * 100 & "%"
        End If
    Next myRange
    
End Sub