Sub 余分な空白を削除()
    
    Dim myRange As Range

    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)
        If myRange.Value <> "" And TypeName(myRange.Value) <> "Date" Then   'セルの値が空白,日付の場合は処理をしない
            myRange.Value = Trim(myRange.Value)
        End If
    Next myRange
    
End Sub