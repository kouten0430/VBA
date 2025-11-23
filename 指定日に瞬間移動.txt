Sub 指定日に瞬間移動()
    '任意の列を垂直方向に検索します
    Dim 日付列 As Long
    Dim 行終 As Long
    Dim 今日 As String
    Dim tmp As String
    
    日付列 = 1
    行終 = 9999
    
    今日 = InputBox("指定日を入力")
    If 今日 = "" Then Exit Sub
    
    今日 = StrConv(今日, vbNarrow)
    
    For i = 1 To 行終
        If TypeName(Cells(i, 日付列).Value) = "Date" Then
            tmp = Day(Cells(i, 日付列).Value)
            
            If 今日 = tmp Then
                ActiveWindow.ScrollRow = Cells(i, 日付列).Row
                ActiveWindow.ScrollColumn = Cells(i, 日付列).Column
                
                Exit For
                
            End If
            
        End If
        
    Next i
    
End Sub