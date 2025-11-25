Sub n列挿入する軽負荷版()
    
    '現在の列の左方向にn列挿入する
    '現在の列の左の書式がコピーされる
    
    Dim n As Long

    n = InputBox("挿入する列数を入力して下さい")
    
    If n >= 1 And n <= 15000 Then
        Range(Columns(ActiveCell.Column), Columns(ActiveCell.Column + n - 1)).Insert xlShiftToRight, xlFormatFromLeftOrAbove
    ElseIf n > 15000 Then
        MsgBox "数値が大きすぎます"
    Else
        MsgBox "1以上の数値を入力して下さい"
    End If
    
End Sub