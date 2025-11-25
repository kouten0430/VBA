Sub n列挿入する式コピー版()
    '現在の列の右方向にn列挿入する
    '現在の列の書式がコピーされる
    '選択セルの式がコピーされる（複数選択可）
    'Selection(1)が現在の列となる
    'Selection(1)と同じ列の選択セルの式をコピーする
    Dim n As Long
    Dim myRange As Range
    
    If Selection.Address = Selection.EntireRow.Address Then
        MsgBox "行全体が選択されています"
        Exit Sub
    ElseIf Selection.Address = Selection.EntireColumn.Address Then
        MsgBox "列全体が選択されています"
        Exit Sub
    End If

    n = InputBox("挿入する列数を入力して下さい")

    If n >= 1 And n <= 15000 Then
        Range(Columns(Selection(1).Column + 1), Columns(Selection(1).Column + 1 + n - 1)).Insert xlShiftToRight, xlFormatFromLeftOrAbove
        
        For Each myRange In Selection
            If myRange.Column = Selection(1).Column Then
                Range(Cells(myRange.Row, myRange.Column + 1), Cells(myRange.Row, myRange.Column + 1 + n - 1)).FormulaR1C1 _
                = Cells(myRange.Row, myRange.Column).FormulaR1C1
            End If
        Next myRange
    ElseIf n > 15000 Then
        MsgBox "数値が大きすぎます"
    Else
        MsgBox "1以上の数値を入力して下さい"
    End If
    
End Sub