Sub 数字を数値に変換()
    
    Dim myRange As Range

    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)
        If myRange.Value <> "" And TypeName(myRange.Value) <> "Date" Then
        'セルの値が空白,日付の場合は処理をしない
        '数字以外はVal関数が0を返すので選択しないで下さい
            myRange.Value = Trim(myRange.Value)
            myRange.Value = Replace(myRange.Value, vbLf, "")
            myRange.Value = Replace(myRange.Value, vbCrLf, "")
            myRange.Value = Replace(myRange.Value, "'", "")
            myRange.Value = Replace(myRange.Value, ",", "")
            myRange.Value = Val(myRange.Value)
        End If
    Next myRange
    
End Sub