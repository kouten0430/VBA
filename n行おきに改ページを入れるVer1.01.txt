Sub n行おきに改ページを入れる()
    '選択範囲の行数がnで割り切れない場合は下端に改ページは入りません
    Dim n As Variant
    Dim y As Long
    Dim ye As Long
    
    n = Application.InputBox(Prompt:="選択範囲内でn行おきに改ページを入れます", Type:=1)
        If TypeName(n) = "Boolean" Or n < 1 Then
            Exit Sub
        End If
    
    y = Selection.Row   '選択範囲の最初の行
    ye = Selection.Rows(Selection.Rows.Count).Row   '選択範囲の最終行
    
    For i = y To ye + 1 Step n
        If i > 1 Then   '1行目以前に改ページを入れることはできない
            ActiveSheet.HPageBreaks.Add (Cells(i, 1))
        End If
    Next i
    
End Sub