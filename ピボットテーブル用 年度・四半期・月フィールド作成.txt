Sub ピボットテーブル用年度四半期月フィールド作成()
    '日付データが入ったフィールドの１セルを選択（どれでも良い）して実行
    Dim X As Long
    Dim YS As Long
    Dim YE As Long
    
    If TypeName(ActiveCell.Value) = "Date" Then
        X = ActiveCell.Column
        YS = ActiveCell.CurrentRegion.Row
        YE = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row

        Range(Columns(ActiveCell.Column), Columns(ActiveCell.Column + 2)).Insert xlShiftToRight, _
        xlFormatFromLeftOrAbove
    
        Cells(YS, X).Value = "年度"
        Range(Cells(YS + 1, X), Cells(YE, X)).FormulaR1C1 = _
        "=IF(RC[3]="""","""",IF(MONTH(RC[3])<=3,YEAR(RC[3])-1&""年度"",YEAR(RC[3])&""年度""))"

        Cells(YS, X + 1).Value = "四半期"
        Range(Cells(YS + 1, X + 1), Cells(YE, X + 1)).FormulaR1C1 = _
        "=IF(RC[2]="""","""",IF(MONTH(RC[2])<=3,""第4四半期"",IF(MONTH(RC[2])<=6,""第1四半期"",IF(MONTH(RC[2])<=9,""第2四半期"",""第3四半期""))))"
    
        Cells(YS, X + 2).Value = "月"
        Range(Cells(YS + 1, X + 2), Cells(YE, X + 2)).FormulaR1C1 = _
        "=IF(RC[1]="""","""",MONTH(RC[1])&""月"")"
        
    Else
        MsgBox "日付データの入ったセルを選択して下さい"
        
    End If
End Sub