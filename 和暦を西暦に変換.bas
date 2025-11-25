Sub 和暦を西暦に変換()
    'シリアル値以外の平成29やH29などの文字列から数値を抜き出し西暦に変換する
    
    Dim myRange As Range
    Dim AD As Long

    For Each myRange In Selection
        For i = 1 To Len(myRange)
            If IsNumeric(Mid(myRange, i, 1)) Then
                AD = AD & Mid(myRange, i, 1)
            End If
        Next i
        
        myRange.Value = AD + 1988
        
        AD = Empty
        
    Next myRange
    
End Sub