Sub 今日のシートに瞬間移動シート名検索版()
    'Sheets(1)から順に検索して、今日と同じシート名のシートをアクティブにします
    Dim 今日 As String
    Dim j As Integer
    Dim tmp As String
    
    今日 = Month(Date) & "月"
    
    For j = 1 To Sheets.Count
        tmp = Sheets(j).Name
            
        If StrConv(今日, vbNarrow) = StrConv(tmp, vbNarrow) Then
            Sheets(j).Activate
                
            Exit Sub
                
        End If
        
    Next j
    
End Sub