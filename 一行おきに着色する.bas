Sub 一行おきに着色する()
    '選択範囲の可視行の偶数行に着色します
    '色は乱数で決まります。好きな色が出るまで実行してみて下さい
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 列始 As Long
    Dim 列終 As Long
    Dim 色 As Long
    Dim cnt As Long
    Dim i As Long
    
    行始 = Selection.Row
    行終 = Selection.Rows(Selection.Rows.Count).Row
    列始 = Selection.Column
    列終 = Selection.Columns(Selection.Columns.Count).Column
    色 = WorksheetFunction.RandBetween(0, 16777215)
    
    cnt = 1
    
    For i = 行始 To 行終
        If Rows(i).Hidden = True Then GoTo continue '非表示行は処理対象としない
        
        If cnt Mod 2 = 0 Then  '可視行を上から数えて偶数の時のみ処理する
            Range(Cells(i, 列始), Cells(i, 列終)).Interior.Color = 色
            
        End If
        
        cnt = cnt + 1
        
continue:

    Next i
    
End Sub