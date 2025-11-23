Sub 代日直の要否を判定()
    'ローカル的なマクロです
    Dim 列 As Long
    Dim 行終 As Long
    Dim 色 As Long
    Dim i As Long
    Dim 待機 As Range
    Dim 日直 As Range
    
    列 = Range("F1").Column
    行終 = Cells(Rows.Count, 列).End(xlUp).Row
    色 = 16751103
    
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    
    For i = 1 To 行終
        If Cells(i, 列).Value Like "自宅待機*" And Format(Cells(i, Range("A1").Column).Value, "aaa") <> "土" And Format(Cells(i, Range("A1").Column).Value, "aaa") <> "日" Then
        
            Set 待機 = Nothing
            Set 日直 = Nothing
        
            On Error Resume Next
            
            Set 待機 = Range(Cells(i, Range("W1").Column), Cells(i, Range("AH1").Column)).Find("◎", LookAt:=xlWhole)
            If Not 待機 Is Nothing Then
                Range(Cells(i + 1, Range("N1").Column), Cells(i + 1, Range("V1").Column)).Interior.Color = 色
                Range(Cells(i + 1, Range("AI1").Column), Cells(i + 1, Range("AT1").Column)).Interior.Color = 色
            End If
            
            Set 日直 = Range(Cells(i, Range("W1").Column), Cells(i, Range("AH1").Column)).Find("○", LookAt:=xlWhole)
            If Not 日直 Is Nothing Then
                Range(Cells(i + 2, Range("N1").Column), Cells(i + 2, Range("V1").Column)).Interior.Color = 色
                Range(Cells(i + 2, Range("AI1").Column), Cells(i + 2, Range("AT1").Column)).Interior.Color = 色
            End If
            
            On Error GoTo 0

        End If
        
    Next i
    
End Sub