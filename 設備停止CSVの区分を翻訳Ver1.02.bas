Sub 設備停止CSVの区分を翻訳()
    '設備停止CSV其の壱～其のｎに組み込んで使用する
    Dim 検索文字 As String
    Dim 壱 As String
    Dim 弐 As String
    Dim 参 As String
    Dim 四 As String
    Dim 伍 As String
    Dim 列 As Long
    Dim 行始 As Long
    Dim 行終 As Long
    Dim i As Long
    
    '---ここから計画区分を翻訳する処理---

    検索文字 = "計画区分"   'フィールド名
    壱 = "計画" '区分1の時の名称
    弐 = "計画外"   '区分2の時の名称
    参 = "計画変更" '区分3の時の名称
    
    列 = Cells.Find(検索文字, LookAt:=xlWhole).Column
    行始 = Cells.Find(検索文字, LookAt:=xlWhole).Row
    行終 = Cells(Rows.Count, 列).End(xlUp).Row
    
    For i = 行始 To 行終
        If Cells(i, 列).Value = 1 Then
            Cells(i, 列).Value = 壱
        ElseIf Cells(i, 列).Value = 2 Then
            Cells(i, 列).Value = 弐
        ElseIf Cells(i, 列).Value = 3 Then
            Cells(i, 列).Value = 参
        End If
    Next i
    
    '以下、条件付き書式の設定（自動記録）
    Columns(列).FormatConditions.Add Type:=xlTextString, String:="計画外", _
        TextOperator:=xlContains
    Columns(列).FormatConditions(Columns(列).FormatConditions.Count).SetFirstPriority
    With Columns(列).FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Columns(列).FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Columns(列).FormatConditions(1).StopIfTrue = False
    Columns(列).FormatConditions.Add Type:=xlTextString, String:="計画変更", _
        TextOperator:=xlContains
    Columns(列).FormatConditions(Columns(列).FormatConditions.Count).SetFirstPriority
    With Columns(列).FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Columns(列).FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Columns(列).FormatConditions(1).StopIfTrue = False
    
    '---ここから要求期間　毎連区分を翻訳する処理---

    検索文字 = "要求期間　毎連区分"   'フィールド名
    壱 = "単日" '区分1の時の名称
    弐 = "毎日"   '区分2の時の名称
    参 = "連続" '区分3の時の名称
    
    列 = Cells.Find(検索文字, LookAt:=xlWhole).Column
    行始 = Cells.Find(検索文字, LookAt:=xlWhole).Row
    行終 = Cells(Rows.Count, 列).End(xlUp).Row
    
    For i = 行始 To 行終
        If Cells(i, 列).Value = 1 Then
            Cells(i, 列).Value = 壱
        ElseIf Cells(i, 列).Value = 2 Then
            Cells(i, 列).Value = 弐
        ElseIf Cells(i, 列).Value = 3 Then
            Cells(i, 列).Value = 参
        End If
    Next i
    
    '以下、条件付き書式の設定（自動記録）
    Columns(列).FormatConditions.Add Type:=xlTextString, String:="連続", _
        TextOperator:=xlContains
    Columns(列).FormatConditions(Columns(列).FormatConditions.Count).SetFirstPriority
    With Columns(列).FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Columns(列).FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Columns(列).FormatConditions(1).StopIfTrue = False
    Columns(列).FormatConditions.Add Type:=xlTextString, String:="毎日", _
        TextOperator:=xlContains
    Columns(列).FormatConditions(Columns(列).FormatConditions.Count).SetFirstPriority
    With Columns(列).FormatConditions(1).Font
        .Color = -16754788
        .TintAndShade = 0
    End With
    With Columns(列).FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 10284031
        .TintAndShade = 0
    End With
    Columns(列).FormatConditions(1).StopIfTrue = False
    
    '---ここから取扱区分を翻訳する処理---

    検索文字 = "取扱区分"   'フィールド名
    壱 = "一般" '区分1の時の名称
    弐 = "協議" '区分2の時の名称
    参 = "主要" '区分3の時の名称
    四 = "ＣＰ一般" '区分4の時の名称
    伍 = "ＣＰ主要" '区分5の時の名称
    
    列 = Cells.Find(検索文字, LookAt:=xlWhole).Column
    行始 = Cells.Find(検索文字, LookAt:=xlWhole).Row
    行終 = Cells(Rows.Count, 列).End(xlUp).Row
    
    For i = 行始 To 行終
        If Cells(i, 列).Value = 1 Then
            Cells(i, 列).Value = 壱
        ElseIf Cells(i, 列).Value = 2 Then
            Cells(i, 列).Value = 弐
        ElseIf Cells(i, 列).Value = 3 Then
            Cells(i, 列).Value = 参
        ElseIf Cells(i, 列).Value = 4 Then
            Cells(i, 列).Value = 四
        ElseIf Cells(i, 列).Value = 5 Then
            Cells(i, 列).Value = 伍
        End If
    Next i
    
End Sub