Function BCD2BIN(Bcd As String) As String
    'ˆø”‚Éw’è‚µ‚½BCD‚ğ2i”‚É•ÏŠ·‚µ‚Ü‚·
    Dim BcdL As Integer
    Dim i As Integer
    
    BcdL = Len(Bcd)
    
    For i = 1 To BcdL
        BCD2BIN = BCD2BIN & WorksheetFunction.Hex2Bin(Mid(Bcd, i, 1), 4)
    Next i
End Function