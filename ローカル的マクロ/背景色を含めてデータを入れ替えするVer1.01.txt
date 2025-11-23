Sub 背景色を含めてデータを入れ替えする()
    'ローカル的なマクロです
    'MergeCellsは指定範囲がすべて結合されていればTrue、すべて結合されていなければFalse、部分的に結合されていればNullとなる
    Dim i As Integer
    Dim 色(1 To 2) As Long
    Dim tmp As String
    
    If Selection.Areas.Count <> 2 Then Exit Sub
    If Selection.Areas(1).Count > 1 And (Selection.Areas(1).MergeCells = False Or IsNull(Selection.Areas(1).MergeCells)) Then Exit Sub
    If Selection.Areas(2).Count > 1 And (Selection.Areas(2).MergeCells = False Or IsNull(Selection.Areas(2).MergeCells)) Then Exit Sub
        
    For i = 1 To 2
        If Selection.Areas(i).Cells(1, 1).Interior.ColorIndex = xlNone Then    '塗りつぶしなしの判定はColorIndexプロパティでのみ可能
            色(i) = xlNone  '塗りつぶしなしの定数
        Else
            色(i) = Selection.Areas(i).Cells(1, 1).Interior.Color
        End If

    Next i

    tmp = Selection.Areas(1).Cells(1, 1).Value
    Selection.Areas(1).Cells(1, 1).Value = Selection.Areas(2).Cells(1, 1).Value
    Selection.Areas(1).Interior.Color = 色(2)
    Selection.Areas(2).Cells(1, 1).Value = tmp
    Selection.Areas(2).Interior.Color = 色(1)
    
End Sub