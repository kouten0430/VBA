Sub 左からn文字目の後ろに改行を入れる()
    '左から数えてn文字目の後ろにセル内改行を入れます
    Dim myRange As Range
    Dim ns As Variant
    
    ns = Application.InputBox(Prompt:="左から何文字目の後ろに改行を入れますか？", Type:=1)
        If TypeName(ns) = "Boolean" Then
            Exit Sub
        End If
    
    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)
        If myRange.Value <> "" And TypeName(myRange.Value) <> "Date" Then   'セルの値が空白,日付の場合は処理をしない
            myRange.Value = Application.WorksheetFunction.Replace(myRange.Value, ns + 1, 0, vbLf)
        End If
    Next myRange
    
End Sub