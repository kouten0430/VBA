Sub ラベルをずべて繰り返す()
    Dim 列 As Long
    Dim 行始 As Long
    Dim 行終 As Long
    Dim i As Long

    列 = Cells.Find("氏名/組織名", LookAt:=xlWhole).Column
    行始 = Cells.Find("氏名/組織名", LookAt:=xlWhole).Row
    行終 = Cells.Find("氏名/組織名", LookAt:=xlWhole).CurrentRegion.Rows(Cells.Find("氏名/組織名", LookAt:=xlWhole).CurrentRegion.Rows.Count).Row
    
    For i = 行始 To 行終
        If Cells(i, 列).Value = "" Then
            Cells(i, 列).Value = Cells(i - 1, 列).Value
            
        End If
        
    Next i

End Sub