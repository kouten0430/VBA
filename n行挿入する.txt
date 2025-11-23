Sub n行挿入する()
    
    '現在の行の上方向にn行挿入する
    '現在の行の上の書式がコピーされる
    
    Dim n As Integer

    n = InputBox("挿入する行数を入力して下さい")
    
    If n >= 1 And n <= 500 Then
        For i = 1 To n
            ActiveCell.EntireRow.Insert xlShiftDown, xlFormatFromLeftOrAbove
        Next i
    ElseIf n > 500 Then
        MsgBox "数値が大きすぎます"
    Else
        MsgBox "1以上の数値を入力して下さい"
        Exit Sub
    End If
    
End Sub