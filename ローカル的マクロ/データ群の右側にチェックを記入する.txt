Sub データ群の右側にチェックを記入する()
'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    Dim 群右側 As Long
    Dim 行 As Long

    If Selection.Areas(1).Cells(1, 1).Value <> "" Then
        V = ChrW(10003)
    
        Set myRange = Selection.Areas(1).Cells(1, 1).CurrentRegion.Find(V, LookAt:=xlWhole)
        
        If myRange Is Nothing Then
            群右側 = Selection.Areas(1).Cells(1, 1).CurrentRegion.Columns(Selection.Areas(1).Cells(1, 1).CurrentRegion.Columns.Count).Column + 1
    
        Else
            群右側 = myRange.Column
    
        End If
    
        For Each myRange In Selection
            If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
                行 = myRange.Row
                Cells(行, 群右側).Value = V
                
            End If
        Next myRange

    End If

End Sub