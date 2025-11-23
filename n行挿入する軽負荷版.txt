Sub n行挿入する軽負荷版()
    
    '現在の行の上方向にn行挿入する
    '現在の行の上の書式がコピーされる
    
    Dim n As Long

    n = InputBox("挿入する行数を入力して下さい")
    
    If n >= 1 And n <= 1000000 Then
        Rows(ActiveCell.Row & ":" & ActiveCell.Row + n - 1).Insert xlShiftDown, xlFormatFromLeftOrAbove
    ElseIf n > 1000000 Then
        MsgBox "数値が大きすぎます"
    Else
        MsgBox "1以上の数値を入力して下さい"
    End If
    
End Sub