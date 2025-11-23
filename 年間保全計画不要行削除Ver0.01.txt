Sub 年間保全計画不要行削除()
    Dim YS As Long
    Dim YE As Long
    
    ActiveCell.Worksheet.AutoFilter.Range.AutoFilter Field:=1, Criteria1:= _
    "電気所等(機能場所)"
    YS = ActiveCell.Worksheet.AutoFilter.Range.Row
    YE = ActiveCell.Worksheet.AutoFilter.Range.Rows(ActiveCell. _
    Worksheet.AutoFilter.Range.Rows.Count).Row

    Rows(YS + 1 & ":" & YE).Select
    Selection.Delete
    ActiveCell.Worksheet.AutoFilter.Range.AutoFilter Field:=1   'Criteria1の省略でAllとなる
End Sub