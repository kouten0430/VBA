Sub 右からn文字目の前に改行を入れる()
    '右から数えてn文字目の前にセル内改行を入れます
    Dim myRange As Range
    Dim nr As Variant
    
    nr = Application.InputBox(Prompt:="右から何文字目の前に改行を入れますか？", Type:=1)
        If TypeName(nr) = "Boolean" Then
            Exit Sub
        End If
    
    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)
        If myRange.Value <> "" And TypeName(myRange.Value) <> "Date" Then   'セルの値が空白,日付の場合は処理をしない
            myRange.Value = Application.WorksheetFunction.Replace(myRange.Value, Len(myRange.Value) - nr + 1, 0, vbLf)
        End If
    Next myRange
    
End Sub