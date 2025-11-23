Sub 今日のシートに瞬間移動()
    '任意の列を垂直方向に検索します
    'Sheets(1)から順に検索して、今日のシリアル値が入っているシートをアクティブにします
    Dim 日付列 As Long
    Dim 行終 As Long
    Dim 今日 As String
    Dim j As Integer
    Dim i As Integer
    Dim tmp As String
    
    日付列 = 1
    行終 = 999
    
    今日 = Year(Date) & Month(Date) & Day(Date)
    
    For j = 1 To Sheets.Count
    
        For i = 1 To 行終
            If TypeName(Sheets(j).Cells(i, 日付列).Value) = "Date" Then
                tmp = Year(Sheets(j).Cells(i, 日付列).Value) & Month(Sheets(j).Cells(i, 日付列).Value) & Day(Sheets(j).Cells(i, 日付列).Value)
            
                If 今日 = tmp Then
                    Sheets(j).Activate
                
                    Exit Sub
                
                End If
            
            End If
        
        Next i
        
    Next j
    
End Sub