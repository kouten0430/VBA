Function BIN2BCD(Bin As String) As String
    '引数に指定した2進数をBCDに変換します
    Dim BinL As Integer
    Dim i As Integer
    
    Bin = "000" & Bin
    BinL = Len(Bin)
    
    For i = 3 To BinL - 1 Step 4
        If 9 >= WorksheetFunction.Bin2Dec(Mid(Bin, BinL - i, 4)) Then
            BIN2BCD = WorksheetFunction.Bin2Dec(Mid(Bin, BinL - i, 4)) & BIN2BCD
        Else    '9を超える場合はエラー表示し、プログラムを終了する
            BIN2BCD = "Error"
            Exit Function
        End If
    Next i
End Function