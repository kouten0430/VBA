Sub 翌日に瞬間移動()
    '現在上端に表示している日の翌日に移動します
    '任意の列を垂直方向に検索します
    Dim 日付列 As Long
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 翌日 As String
    Dim i As Long
    Dim tmp As String
    
    日付列 = 1
    行始 = ActiveWindow.VisibleRange.Item(1, 1).Row
    行終 = 9999
    
    If TypeName(Cells(行始, 日付列).Value) = "Date" Then
        翌日 = Year(Cells(行始, 日付列).Value + 1) & Month(Cells(行始, 日付列).Value + 1) & Day(Cells(行始, 日付列).Value + 1)
        
        For i = 行始 To 行終
            If TypeName(Cells(i, 日付列).Value) = "Date" Then
                tmp = Year(Cells(i, 日付列).Value) & Month(Cells(i, 日付列).Value) & Day(Cells(i, 日付列).Value)
                
                If 翌日 = tmp Then
                    ActiveWindow.ScrollRow = Cells(i, 日付列).Row
                    ActiveWindow.ScrollColumn = Cells(i, 日付列).Column
                    
                    Exit For
                    
                End If
                
            End If
            
        Next i
        
    End If
    
End Sub