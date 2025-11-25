Sub 西暦を和暦に変換()
    'シリアル値以外の2017や2017年などの文字列から数値を抜き出し和暦に変換する
    
    Dim myRange As Range
    Dim JC As Long

    For Each myRange In Selection
        For i = 1 To Len(myRange)
            If IsNumeric(Mid(myRange, i, 1)) Then
                JC = JC & Mid(myRange, i, 1)
            End If
        Next i
        
        myRange.Value = "H" & JC - 1988
        
        JC = Empty
        
    Next myRange
    
End Sub