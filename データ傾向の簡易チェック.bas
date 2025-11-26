Sub ƒf[ƒ^ŒXŒü‚ÌŠÈˆÕƒ`ƒFƒbƒN()
    Dim ‹}•Ï’l As Double
    Dim ¬ As Integer
    Dim ‹}‰º~ As Integer
    Dim ‘å As Integer
    Dim ‹}ã¸ As Integer
    Dim “¯ As Integer
    
    ‹}•Ï’l = 3
    ã¸‹–—e = 1
    
    If Not IsNumeric(Selection(1)) Then
        MsgBox "”’lˆÈŠO‚ªŠÜ‚Ü‚ê‚Ä‚¢‚Ü‚·!"
        Exit Sub
    End If

    For i = 2 To Selection.Count
        If IsNumeric(Selection(i).Value) Then
            If Selection(i - 1).Value > Selection(i).Value Then
                ¬ = ¬ + 1
                If Selection(i - 1).Value - Selection(i).Value >= ‹}•Ï’l Then
                    ‹}‰º~ = ‹}‰º~ + 1
                End If
            ElseIf Selection(i - 1).Value < Selection(i).Value Then
                ‘å = ‘å + 1
                If Selection(i).Value - Selection(i - 1).Value >= ‹}•Ï’l Then
                    ‹}ã¸ = ‹}ã¸ + 1
                End If
            Else
                “¯ = “¯ + 1
            End If
        Else
            MsgBox "”’lˆÈŠO‚ªŠÜ‚Ü‚ê‚Ä‚¢‚Ü‚·!"
            Exit Sub
        End If
    Next i

    If “¯ = Selection.Count - 1 Then
        MsgBox "‘S‚­•Ï‰»‚È‚µ"
    ElseIf ¬ = (Selection.Count - 1) Then
        MsgBox "‚¸‚Á‚Æ‰º‚ª‚èŒXŒü" & vbCrLf & "‹}‰º~" & ‹}‰º~ & "‰ñ‚ ‚è"
    ElseIf ‘å = 0 Then
        MsgBox "‚ä‚é‚â‚©‚É‰º‚ª‚èŒXŒü" & vbCrLf & "‹}‰º~" & ‹}‰º~ & "‰ñ‚ ‚è"
    ElseIf ‘å > 0 And ‘å <= ã¸‹–—e And ‹}ã¸ = 0 And ¬ >= 2 Then
        MsgBox "‚ä‚é‚â‚©‚É‰º‚ª‚èŒXŒüiã¸ " & ‘å & "‰ñj" & vbCrLf & "‹}‰º~" & ‹}‰º~ & "‰ñ‚ ‚è"
    Else
        MsgBox "‹}‰º~" & ‹}‰º~ & "‰ñ‚ ‚è" & vbCrLf & "‹}ã¸" & ‹}ã¸ & "‰ñ‚ ‚è" & vbCrLf
    End If

End Sub