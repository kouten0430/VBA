Sub 指定日に水平移動()
    '任意の行を水平方向に検索します
    Dim 日付行 As Long
    Dim 列終 As Long
    Dim 指定日 As String
    Dim i As Integer
    Dim tmp As String
    
    日付行 = 6
    列終 = 999
    
    指定日 = InputBox("指定日を入力")
    If 指定日 = "" Then Exit Sub
    
    指定日 = StrConv(指定日, vbNarrow)
    
    For i = 1 To 列終
        If TypeName(Cells(日付行, i).Value) = "Date" Then
            tmp = Day(Cells(日付行, i).Value)
            
            If 指定日 = tmp Then
                ActiveWindow.ScrollColumn = i
                
                Exit For
                
            End If
            
        End If
        
    Next i
    
End Sub