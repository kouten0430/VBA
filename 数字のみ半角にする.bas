Sub ”š‚Ì‚İ”¼Šp‚É‚·‚é()
    '‘I‘ğ”ÍˆÍ‚É‘Î‚µ‚Äˆ—‚ğs‚¢‚Ü‚·
    Dim myRange As Range
    Dim i As Integer

    For Each myRange In Selection
        For i = 1 To Len(myRange.Value)
            If Mid(myRange.Value, i, 1) Like "[‚O-‚X]" Then
                myRange.Value = Application.WorksheetFunction.Replace(myRange.Value, i, 1, _
                StrConv(Mid(myRange.Value, i, 1), vbNarrow))
            End If
        Next i
    Next myRange
    
End Sub