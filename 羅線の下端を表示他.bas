Sub 羅線の下端を表示()
    Dim x As Integer
    Dim y As Long
    Dim YE As Long
    Dim EL As Integer
    Dim ER As Integer
    Dim EB As Integer
    
    x = ActiveCell.Column   '選択中のセルの列番号を取得する
    y = ActiveCell.Row  '選択中のセルの行番号を取得する
    
    YE = y

    EL = Cells(y, x).Borders(xlEdgeLeft).LineStyle  'セルの左端の罫線の種類を取得する
    ER = Cells(y, x).Borders(xlEdgeRight).LineStyle 'セルの右端の罫線の種類を取得する
    EB = Cells(y, x).Borders(xlEdgeBottom).LineStyle 'セルの下端の罫線の種類を取得する

    Do While (Not EL = xlLineStyleNone Or Not ER = xlLineStyleNone Or Not EB = xlLineStyleNone) 'セルの右か左か下に罫線がある場合は処理を行う
    
        YE = YE + 1 '１行下に進む
    
        EL = Cells(YE, x).Borders(xlEdgeLeft).LineStyle
        ER = Cells(YE, x).Borders(xlEdgeRight).LineStyle
        EB = Cells(YE, x).Borders(xlEdgeBottom).LineStyle

    Loop

    If Not y = YE Then  'ループに一度も入らなかった場合、処理を行わない
        ActiveWindow.ScrollRow = YE + 1
        ActiveWindow.LargeScroll Up:=1
    End If

End Sub
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
    End If

End Sub
Sub 一画面下を表示()
    ActiveWindow.LargeScroll Down:=1
End Sub
Sub 一画面右を表示()
    ActiveWindow.LargeScroll ToRight:=1
End Sub
Sub 一画面上を表示()
    ActiveWindow.LargeScroll Up:=1
End Sub
Sub 一画面左を表示()
    ActiveWindow.LargeScroll ToLeft:=1
End Sub
Sub 上端を表示()
    ActiveWindow.ScrollRow = 1
End Sub
Sub 左端を表示()
    ActiveWindow.ScrollColumn = 1
End Sub