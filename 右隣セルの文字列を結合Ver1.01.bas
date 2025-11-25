Sub 右隣の文字を連結()
    '選択したセルの文字列に右隣nセルの文字列を連結する
    '非表示の列の文字列は連結しません（連結させたくない列を非表示にしておくと便利です）
    Dim Y As Integer
    Dim X As Integer
    Dim XR As Variant
    Dim XC As Variant
    Dim i As Integer
    Dim myRange As Range
    
    XR = Application.InputBox(Prompt:="選択中のセルに右側何セル分の文字を連結しますか？", Type:=1)
        If TypeName(XR) = "Boolean" Then
            Exit Sub
        End If

    XC = Application.InputBox(Prompt:="連結間に挿入する文字を入力して下さい。（ブランクでも可）" _
    & vbCrLf & "改行にする場合はキャンセルして下さい。", Type:=2)
        If TypeName(XC) = "Boolean" Then
            XC = vbLf
        End If

    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)
        Y = myRange.Row
        X = myRange.Column + 1
        i = 1
        Do While i <= XR
            If Columns(X).Hidden = False Then '非表示の列は処理を行わない
                myRange.Value = myRange.Value & XC & Cells(Y, X).Value
                X = X + 1
                i = i + 1
            Else
                X = X + 1
            End If
        Loop
    Next myRange
    
End Sub