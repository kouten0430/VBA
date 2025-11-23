Sub ううう二つ目に選択したセルへ値をもっていく移動版待機日直用()
    'ローカル的なマクロです
    Dim tmp1 As String
    Dim tmp2 As String
    
    If Selection.Count <> 2 Or Selection.Areas(1).Value = "" Then Exit Sub
    
    tmp1 = Selection.Areas(1).Value
    tmp2 = Selection.Areas(1).Offset(0, 1).Value
    
    Selection.Areas(1).ClearContents
    Selection.Areas(1).Offset(0, 1).ClearContents
    
    Selection.Areas(2).Value = tmp1
    Selection.Areas(2).Offset(0, 1).Value = tmp2
    
End Sub