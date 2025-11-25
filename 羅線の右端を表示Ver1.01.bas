Sub 羅線の右端を表示()
    Dim x As Integer
    Dim y As Long
    Dim xe As Integer
    Dim ER As Integer
    Dim EB As Integer
    
    x = ActiveCell.Column   '選択中のセルの列番号を取得する
    y = ActiveCell.Row  '選択中のセルの行番号を取得する
    
    xe = x

    ER = Cells(y, x).Borders(xlEdgeRight).LineStyle 'セルの右端の罫線の種類を取得する
    EB = Cells(y, x).Borders(xlEdgeBottom).LineStyle 'セルの下端の罫線の種類を取得する
    
    Do While (Not ER = xlLineStyleNone Or Not EB = xlLineStyleNone) 'セルの右か下に罫線がある場合は処理を行う
    
        xe = xe + 1 '１列右に進む
    
        ER = Cells(y, xe).Borders(xlEdgeRight).LineStyle
        EB = Cells(y, xe).Borders(xlEdgeBottom).LineStyle

    Loop

    If Not x = xe Then 'ループに一度も入らなかった場合、処理を行わない
        ActiveWindow.ScrollColumn = xe + 1
        ActiveWindow.LargeScroll ToLeft:=1
        
    ElseIf ActiveCell.Address <> ActiveCell.CurrentRegion.Address Then  '罫線が無ければデータ群の右端を表示
        ActiveWindow.ScrollColumn = ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column + 2
        ActiveWindow.LargeScroll ToLeft:=1
        
    End If

End Sub