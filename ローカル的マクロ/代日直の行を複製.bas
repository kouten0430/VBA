Sub 代日直の行を複製()
    'ローカル的なマクロです
    Dim 列 As Long
    Dim 行終 As Long
    Dim i As Long
    
    列 = Range("F1").Column
    行終 = Cells(Rows.Count, 列).End(xlUp).Row
    
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    
    For i = 行終 To 1 Step -1
        If Cells(i, 列).Value = "代日直" Then
            Cells(i, 列 - 2).Value = "'17-20"
            Cells(i, 列 - 1).Value = "'18-20"
            
            Rows(i).Insert xlShiftDown, xlFormatFromRightOrBelow
            
            Range(Cells(i, Range("A1").Column), Cells(i, Range("B1").Column)).FormulaR1C1 = Range(Cells(i + 1, Range("A1").Column), Cells(i + 1, Range("B1").Column)).FormulaR1C1
            Range(Cells(i, Range("F1").Column), Cells(i, Range("F1").Column)).FormulaR1C1 = Range(Cells(i + 1, Range("F1").Column), Cells(i + 1, Range("F1").Column)).FormulaR1C1
            Range(Cells(i, Range("BF1").Column), Cells(i, Range("ZF1").Column)).FormulaR1C1 = Range(Cells(i + 1, Range("BF1").Column), Cells(i + 1, Range("ZF1").Column)).FormulaR1C1
            
            Cells(i, 列 - 2).Value = "'8-00"
            Cells(i, 列 - 1).Value = "'8-40"

        End If
        
    Next i
    
End Sub