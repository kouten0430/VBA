Sub 経年を生成()
    Dim 年度 As Variant
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 列始 As Long
    Dim 列終 As Long
    Dim 設備運開日 As Long
    Dim 製造年 As Long
    Dim tmp As Long
    
    年度 = Application.InputBox(prompt:="経年を算出する年度を入力して下さい。（西暦で入力）", Type:=1)
    If TypeName(年度) = "Boolean" Then Exit Sub
    
    '---ここから行番号・列番号を取得する処理---
    
    行始 = ActiveCell.CurrentRegion.Row
    行終 = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row
    列始 = ActiveCell.CurrentRegion.Column
    列終 = ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column
    
    For i = 列始 To 列終
        If Cells(行始, i).Value = "設備運開日" Then
            設備運開日 = i
        ElseIf Cells(行始, i).Value = "製造年月日" Then
            製造年 = i
        End If
    Next i
        
    If 設備運開日 = 0 Then 設備運開日 = 9999
    If 製造年 = 0 Then 製造年 = 9999
    
    '---ここから経年を生成する処理---
    
    Cells(行始, 列終 + 1).Value = "経年（" & 年度 & "年度時点）"
    
    For i = 行始 + 1 To 行終
    
        If Cells(i, 製造年).Value <> "" Then    '製造年が空白なら運開年から生成する
            tmp = 製造年
        Else
            tmp = 設備運開日
        End If
    
        If Cells(i, tmp).Value <> "" Then
            If Month(Cells(i, tmp).Value) < 4 Then
                Cells(i, 列終 + 1).Value = 年度 - (Year(Cells(i, tmp).Value) - 1)
            Else
                Cells(i, 列終 + 1).Value = 年度 - Year(Cells(i, tmp).Value)
            End If
        End If
    Next i
    
End Sub