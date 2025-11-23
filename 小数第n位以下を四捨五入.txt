Sub 小数第n位以下を四捨五入()
    Dim n As Variant
    Dim myRange As Range
    n = Application.InputBox(Prompt:="小数第何位以下を四捨五入しますか？", Type:=1)
        If TypeName(n) = "Boolean" Then
            Exit Sub
        End If
    For Each myRange In Selection
        If myRange.Value <> 0 And myRange.Value <> "" And _
        TypeName(myRange.Value) <> "String" And TypeName(myRange.Value) <> "Date" Then
        'セルの値が0,空白,文字列,日付のいづれかの場合は処理をしない
            myRange.Value = Round(myRange.Value, n - 1)  '小数第n位以下は四捨五入
        End If
    Next myRange
    
End Sub