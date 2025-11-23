Sub 行の高さや列幅を他の行または列に適用する()
    '適用先の行全体または列全体を選択した状態で実行する
    'インプットボックスで適用元の行または列を指定する
    Dim 列番号 As String
    Dim 列幅 As Double
    Dim 行番号 As Long
    Dim 行高さ As Double
    
    If Selection.Address = Selection.EntireColumn.Address Then
        列番号 = InputBox("列幅の適用元となる列番号をアルファベットで指定")
        列番号 = StrConv(列番号, vbNarrow)
        
        If 列番号 Like "*[!A-Za-z]*" Then GoTo エラー処理
        
        On Error GoTo エラー処理
            列幅 = Columns(列番号).ColumnWidth
        On Error GoTo 0
        
        Selection.ColumnWidth = 列幅
        
    ElseIf Selection.Address = Selection.EntireRow.Address Then
        行番号 = InputBox("行高さの適用元となる行番号を指定")
        
        On Error GoTo エラー処理
            行高さ = Rows(行番号).RowHeight
        On Error GoTo 0
        
        Selection.RowHeight = 行高さ
    
    End If
    
    Exit Sub

エラー処理:
    MsgBox "存在しない行または列です。"
End Sub