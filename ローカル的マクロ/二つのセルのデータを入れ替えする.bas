Sub 二つのセルのデータを入れ替えする()
    'ローカル的なマクロです
    'MergeCellsは指定範囲がすべて結合されていればTrue、すべて結合されていなければFalse、部分的に結合されていればNullとなる
    Dim tmp As String
    
    If Selection.Areas.Count <> 2 Then Exit Sub
    If Selection.Areas(1).Count > 1 And (Selection.Areas(1).MergeCells = False Or IsNull(Selection.Areas(1).MergeCells)) Then Exit Sub
    If Selection.Areas(2).Count > 1 And (Selection.Areas(2).MergeCells = False Or IsNull(Selection.Areas(2).MergeCells)) Then Exit Sub

    tmp = Selection.Areas(1).Cells(1, 1).Value
    Selection.Areas(1).Cells(1, 1).Value = Selection.Areas(2).Cells(1, 1).Value
    Selection.Areas(2).Cells(1, 1).Value = tmp
    
End Sub