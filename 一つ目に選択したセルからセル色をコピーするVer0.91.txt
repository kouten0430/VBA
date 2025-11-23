Sub 一つ目に選択したセルからセル色をコピーする()
    '二つ目以降に選択したセルへセル色をコピーする
    Dim i As Integer
    
    If Selection.Areas(1).Count > 1 And (Selection.Areas(1).MergeCells = False Or IsNull(Selection.Areas(1).MergeCells)) Then Exit Sub

    For i = 1 To Selection.Areas.Count
        Selection.Areas(i).Interior.Color = Selection.Areas(1).Cells(1, 1).DisplayFormat.Interior.Color
        
    Next i
    
End Sub