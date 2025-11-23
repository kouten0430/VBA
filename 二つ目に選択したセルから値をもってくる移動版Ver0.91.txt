Sub 二つ目に選択したセルから値をもってくる移動版()
    If Selection.Areas.Count <> 2 Then Exit Sub
    If Selection.Areas(1).Count > 1 And (Selection.Areas(1).MergeCells = False Or IsNull(Selection.Areas(1).MergeCells)) Then Exit Sub
    If Selection.Areas(2).Count > 1 And (Selection.Areas(2).MergeCells = False Or IsNull(Selection.Areas(2).MergeCells)) Then Exit Sub

    Selection.Areas(1).Cells(1, 1).Value = Selection.Areas(2).Cells(1, 1).Value
    Selection.Areas(2).ClearContents
    
End Sub