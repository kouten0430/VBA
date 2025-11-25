Sub 指定位置に瞬間移動アルファベットで指定版()
    Dim y As Variant
    Dim x As String
    Dim flag As Boolean
    
    y = InputBox("表示する「行」を数値で入力" & vbCrLf & "（行はこのままで良い場合、空白 or キャンセル）")
    
    If y = "" Then
        y = ActiveWindow.ScrollRow
    ElseIf y > 1048576 Then
        y = 1048576
        flag = True
    ElseIf y < 1 Then
        y = 1
    End If
    
retry:
    x = InputBox("表示する「列」をアルファベットで入力" & vbCrLf & "（列はこのままで良い場合、空白 or キャンセル）")
    
    If x <> "" Then
        x = StrConv(x, vbNarrow)
        x = StrConv(x, vbUpperCase)
        If x Like "*[!A-Z]*" Then
            MsgBox "列はアルファベットのみ入力可"
            GoTo retry
        End If
        On Error GoTo ErrorHandler
        ActiveWindow.ScrollColumn = Range(x & "1").Column
        On Error GoTo 0
    End If

    ActiveWindow.ScrollRow = y
    If flag Then MsgBox "いしのなかにいる！"
    
    Exit Sub
    
ErrorHandler:
    x = "XFD"
    flag = True
    Resume
    
End Sub