Sub n文字目から指定文字数を別の文字列に置換する()
    
    Dim myRange As Range
    Dim ns As Integer
    Dim ne As Integer
    Dim tm As String
    
    ns = InputBox("何文字目から置換しますか？")
    ne = InputBox("何文字置換しますか？（挿入する場合は0）")
    tm = InputBox("置換後の文字列は？")

    For Each myRange In Selection
        If myRange.Value <> "" And TypeName(myRange.Value) <> "Date" Then   'セルの値が空白,日付の場合は処理をしない
            myRange.Value = Application.WorksheetFunction.Replace(myRange.Value, ns, ne, tm)
        End If
    Next myRange
    
End Sub