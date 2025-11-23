Sub 小数第n位以下を切り上げ()
    Dim n As Variant
    Dim myRange As Range
    n = Application.InputBox(Prompt:="小数第何位以下を切り上げしますか？", Type:=1)
        If TypeName(n) = "Boolean" Then
            Exit Sub
        End If
    For Each myRange In Selection
        If myRange.Value <> 0 And myRange.Value <> "" And _
        TypeName(myRange.Value) <> "String" And TypeName(myRange.Value) <> "Date" Then
        'セルの値が0,空白,文字列,日付のいづれかの場合は処理をしない
            myRange.Value = Application.RoundUp(myRange.Value, n - 1)  '小数第n位以下は切り上げ
        End If
    Next myRange
    
End Sub