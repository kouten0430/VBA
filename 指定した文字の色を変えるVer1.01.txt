Sub 指定した文字の色を変える()
    Dim sm As Variant
    Dim ci As Variant
    Dim r As Range
    Dim i As Integer

smr:
    sm = Application.InputBox(Prompt:="色を変える文字を指定して下さい", Type:=2)
        If TypeName(sm) = "Boolean" Then
            Exit Sub
        ElseIf sm = "" Then
            GoTo smr
        End If
cir:
    ci = Application.InputBox(Prompt:="色を選んで下さい" _
    & vbCrLf & "1:黒" & vbCrLf & "2:白" & vbCrLf & "3:赤" & vbCrLf & _
    "4:明るい緑" & vbCrLf & "5:青" & vbCrLf & "6:黄色" & vbCrLf & _
    "7:ピンク" & vbCrLf & "8:水色" & vbCrLf & "9:明るい赤" & vbCrLf & _
    "10:緑" & vbCrLf & "(11以降の色番号はVBAのヘルプ等で確認下さい)", Type:=1)
        If TypeName(ci) = "Boolean" Then
            Exit Sub
        ElseIf ci < 1 Or ci > 56 Then
            MsgBox "1～56の数値で入力して下さい"
            GoTo cir
        End If

    For Each r In Selection.SpecialCells(xlCellTypeVisible)
    
        i = 1
    
        Do While i <= Len(r)
            If InStr(i, r, sm) > 0 Then
                r.Characters(InStr(i, r, sm), Len(sm)) _
                .Font.ColorIndex = ci
            
                i = InStr(i, r, sm) + Len(sm)
            Else
                Exit Do '永久ループを回避
            End If
        Loop
    Next r
End Sub