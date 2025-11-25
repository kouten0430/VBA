Sub 羅線のあるセルを下端まで選択する()
    Dim x As Integer
    Dim xe As Integer
    Dim y As Long
    Dim ye As Long
    Dim EL As Integer
    Dim ER As Integer
    Dim EB As Integer
    
    x = ActiveCell.Column   '選択中のセルの列番号を取得する
    y = ActiveCell.Row  '選択中のセルの行番号を取得する
    
    xe = Selection.Columns(Selection.Columns.Count).Column
    ye = y

    EL = Cells(y, x).Borders(xlEdgeLeft).LineStyle  'セルの左端の罫線の種類を取得する
    ER = Cells(y, x).Borders(xlEdgeRight).LineStyle 'セルの右端の罫線の種類を取得する
    EB = Cells(y, x).Borders(xlEdgeBottom).LineStyle 'セルの下端の罫線の種類を取得する

    Do While (Not EL = xlLineStyleNone Or Not ER = xlLineStyleNone Or Not EB = xlLineStyleNone) 'セルの右か左か下に罫線がある場合は処理を行う
    
        ye = ye + 1 '１行下に進む
    
        EL = Cells(ye, x).Borders(xlEdgeLeft).LineStyle
        ER = Cells(ye, x).Borders(xlEdgeRight).LineStyle
        EB = Cells(ye, x).Borders(xlEdgeBottom).LineStyle

    Loop

    If Not y = ye Then  'ループに一度も入らなかった場合、処理を行わない
        Range(Cells(y, x), Cells(ye - 1, xe)).Select '選択中のセルから下方向に、罫線があるセルをすべて選択する
    End If

End Sub