Sub ‘I‘ğ”ÍˆÍ‚Ì•¶š‚ğ”¼Šp‚É‚·‚é()
    
    Dim myRange As Range

    For Each myRange In Selection
        myRange.Value = Replace(StrConv(myRange.Value, vbNarrow), "~", "`")
    Next myRange
    
End Sub