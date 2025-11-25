Sub 今日の予定を印刷する()
    '任意の列を垂直方向に検索します
    Dim 日付列 As Long
    Dim 行終 As Long
    Dim 今日 As String
    Dim 明日 As String
    Dim i As Long
    Dim tmp As String
    Dim 上 As Long
    Dim j As Long
    Dim 下 As Long
    
    日付列 = 1
    行終 = Cells(Rows.Count, 日付列).End(xlUp).Row
    
    今日 = Year(Date) & Month(Date) & Day(Date)
    明日 = Year(Date + 1) & Month(Date + 1) & Day(Date + 1)
    
    For i = 1 To 行終
        If TypeName(Cells(i, 日付列).Value) = "Date" Then
            tmp = Year(Cells(i, 日付列).Value) & Month(Cells(i, 日付列).Value) & Day(Cells(i, 日付列).Value)
            
            If 今日 = tmp Then
                上 = Cells(i, 日付列).Row
                
                Exit For
                
            End If
            
        End If
        
    Next i
    
    For j = 上 + 1 To 行終
        If TypeName(Cells(j, 日付列).Value) = "Date" Then
            tmp = Year(Cells(j, 日付列).Value) & Month(Cells(j, 日付列).Value) & Day(Cells(j, 日付列).Value)
            
            If 明日 = tmp Then
                下 = Cells(j, 日付列).Row - 1
                
                GoTo skip
                
            End If
            
        End If
        
    Next j
    
    下 = Cells(行終, 日付列).Row - 1
    
skip:

    With ActiveSheet.PageSetup
        .Orientation = xlLandscape '印刷向き（横）
        .PaperSize = xlPaperA4     '用紙（A4）
        .Zoom = False '拡大縮小（しない：FitToPagesに委ねる）
        .FitToPagesWide = 1 'すべての列をnページに印刷
        .FitToPagesTall = False 'すべての行をnページに印刷（指定しない）

    End With

    Range(Cells(上, "A"), Cells(下, "BE")).PrintOut
    
End Sub