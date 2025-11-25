Sub 末尾に※休工を追加する()
    'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "※休工"
        
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
Sub 末尾に※立会省略を追加する()
    'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "※立会省略"
        
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
Sub 末尾に※ポイント立会を追加する()
    'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "※ポイント立会"
        
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
Sub 末尾に※中止を追加する()
    'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "※中止"
        
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
Sub 末尾に※延期を追加する()
    'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "※延期"
        
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