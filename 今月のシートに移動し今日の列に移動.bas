Sub 今月のシートに移動し今日の列に移動()

    '---ここから今月のシートに移動する処理---
    
    Dim 今月 As String
    Dim j As Integer
    Dim シート名 As String
    
    今月 = Month(Date) & "月"
    
    For j = 1 To Sheets.Count
        シート名 = Sheets(j).Name
            
        If StrConv(今月, vbNarrow) = StrConv(シート名, vbNarrow) Then
            Sheets(j).Activate
                
            Exit For
                
        End If
        
    Next j
    
    '---ここから今日の列に移動する処理---
    
    Dim 日付行 As Long
    Dim 列終 As Long
    Dim 今日 As String
    Dim i As Integer
    Dim tmp As String
    
    日付行 = 3
    列終 = 999
    
    今日 = Year(Date) & Month(Date) & Day(Date)
    
    For i = 1 To 列終
        If TypeName(Cells(日付行, i).Value) = "Date" Then
            tmp = Year(Cells(日付行, i).Value) & Month(Cells(日付行, i).Value) & Day(Cells(日付行, i).Value)
            
            If 今日 = tmp Then
                ActiveWindow.ScrollColumn = i
                
                Exit For
                
            End If
            
        End If
        
    Next i
    
End Sub