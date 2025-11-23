Sub ‹L†‚Ì‚İ”¼Šp‚É‚·‚é()
    '‘I‘ğ”ÍˆÍ‚É‘Î‚µ‚Äˆ—‚ğs‚¢‚Ü‚·
    Dim myRange As Range
    Dim i As Integer

    For Each myRange In Selection
    
        i = 1
    
        Do While i <= Len(myRange.Value)
            If Mid(myRange.Value, i, 1) Like "[!0-9‚O-‚XA-Za-z‚`-‚y‚-‚šƒ@-ƒ–¦-ß]" Then
                myRange.Value = Application.WorksheetFunction.Replace(myRange.Value, i, 1, _
                StrConv(Mid(myRange.Value, i, 1), vbNarrow))
            End If
            
            i = i + 1
            
        Loop
    Next myRange
    
End Sub