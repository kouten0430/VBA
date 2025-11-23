Sub 羅線のあるセルを右端まで選択する()
    Dim x As Integer
    Dim y As Long
    Dim xe As Integer
    Dim ye As Long
    Dim ER As Integer
    Dim EB As Integer
    
    x = ActiveCell.Column   '選択中のセルの列番号を取得する
    y = ActiveCell.Row  '選択中のセルの行番号を取得する
    
    xe = x
    ye = Selection.Rows(Selection.Rows.Count).Row

    ER = Cells(y, x).Borders(xlEdgeRight).LineStyle 'セルの右端の罫線の種類を取得する
    EB = Cells(y, x).Borders(xlEdgeBottom).LineStyle 'セルの下端の罫線の種類を取得する
    
    Do While (Not ER = xlLineStyleNone Or Not EB = xlLineStyleNone) 'セルの右か下に罫線がある場合は処理を行う
    
        xe = xe + 1 '１列右に進む
    
        ER = Cells(y, xe).Borders(xlEdgeRight).LineStyle
        EB = Cells(y, xe).Borders(xlEdgeBottom).LineStyle

    Loop

    If Not x = xe Then 'ループに一度も入らなかった場合、処理を行わない
        Range(Cells(y, x), Cells(ye, xe - 1)).Select '選択中のセルから右方向に、罫線があるセルをすべて選択する
    End If

End Sub