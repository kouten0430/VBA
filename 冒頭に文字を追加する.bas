Sub 冒頭に文字を追加する()
    Dim V As Variant
    Dim myRange As Range
    
    V = Application.InputBox(Prompt:="冒頭に追加する文字を入力して下さい", Type:=2)
        If TypeName(V) = "Boolean" Then
            Exit Sub
        End If
    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
            myRange.Value = V & myRange.Value
        End If
    Next myRange

End Sub