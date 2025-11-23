Sub 年度半期四半期月フィールド作成()
    '日付データが入ったフィールドの１セルを選択（どれでも良い）して実行
    Dim X As Long
    Dim YS As Long
    Dim YE As Long
    
    If TypeName(ActiveCell.Value) = "Date" Then
        X = ActiveCell.Column
        YS = ActiveCell.CurrentRegion.Row
        YE = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row

        Range(Columns(ActiveCell.Column), Columns(ActiveCell.Column + 3)).Insert xlShiftToRight, _
        xlFormatFromLeftOrAbove
    
        Cells(YS, X).Value = "年度"
        Range(Cells(YS + 1, X), Cells(YE, X)).FormulaR1C1 = _
        "=IF(RC[4]="""","""",IF(MONTH(RC[4])<=3,YEAR(RC[4])-1&""年度"",YEAR(RC[4])&""年度""))"
        
        Cells(YS, X + 1).Value = "半期"
        Range(Cells(YS + 1, X + 1), Cells(YE, X + 1)).FormulaR1C1 = _
        "=IF(RC[3]="""","""",IF(AND(MONTH(RC[3])>=4,MONTH(RC[3])<=9),""1H"",""2H""))"

        Cells(YS, X + 2).Value = "四半期"
        Range(Cells(YS + 1, X + 2), Cells(YE, X + 2)).FormulaR1C1 = _
        "=IF(RC[2]="""","""",IF(MONTH(RC[2])<=3,""4Q"",IF(MONTH(RC[2])<=6,""1Q"",IF(MONTH(RC[2])<=9,""2Q"",""3Q""))))"
    
        Cells(YS, X + 3).Value = "月"
        Range(Cells(YS + 1, X + 3), Cells(YE, X + 3)).FormulaR1C1 = _
        "=IF(RC[1]="""","""",MONTH(RC[1])&""月"")"
        
    Else
        MsgBox "日付データの入ったセルを選択して下さい"
        
    End If
End Sub