Sub 末尾に台風特巡を追加する()
    'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "台風特巡"
        
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルにのみ処理を行う
            If myRange.Value <> "" Then 'セルが空白でなければ改行し、空白であれば改行しない
                myRange.Value = myRange.Value & vbLf & V
            Else
                myRange.Value = V
            End If
        End If
        
    Next myRange
    
End Sub