Sub ‰pš‚Ì‚İ‘SŠp‚É‚·‚é()
    '‘I‘ğ”ÍˆÍ‚É‘Î‚µ‚Äˆ—‚ğs‚¢‚Ü‚·
    Dim myRange As Range
    Dim i As Integer

    For Each myRange In Selection
        For i = 1 To Len(myRange.Value)
            If Mid(myRange.Value, i, 1) Like "[A-Za-z]" Then
                myRange.Value = Application.WorksheetFunction.Replace(myRange.Value, i, 1, _
                StrConv(Mid(myRange.Value, i, 1), vbWide))
            End If
        Next i
    Next myRange
    
End Sub