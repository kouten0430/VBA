Sub クリップボードのデータを可視セルのみに貼り付ける()

    'Microsoft Forms 2.0 Object Library に参照設定要
    '複数セルを選択した状態で実行するとアクティブセル（白抜き）が貼り付け開始位置となる
    '結合セルを選択した状態で実行すると結合セルの左上セルが貼り付け開始位置となる
    '結合セルに貼り付けたい場合は結合セルを選択する。もしくは結合セルの左端と同列のセルを選択する
    '改行区切りで下方向に貼り付け
    'Tab区切りがあれば右方向に貼り付け（右方向の非表示セルには未対応）

    Dim Dobj As DataObject
    Dim V As Variant    'クリップボードのデータ全体
    Dim A As Variant    'その内の一行
    Dim rngDest As Range
    Dim R As Range
    Dim i As Integer
    Dim XS As Integer
    Dim XP As Integer
    Dim YS As Integer
    Dim YP As Integer
    Dim YE As Integer

    If ActiveSheet.AutoFilter Is Nothing Then
        MsgBox "オートフィルターが設定されていません！"
        Exit Sub
    End If

    YP = ActiveCell.Row '現在選択しているセルの行番号を取得
    XP = ActiveCell.Column  '現在選択しているセルの列番号を取得
    XP = XP + 1 'AutoFilter.Rangeの左端の列番号が相対的に1となるようにする

    YS = ActiveCell.Worksheet.AutoFilter.Range.Row 'オートフィルターが適用される範囲の上端の行番号を取得
    XS = ActiveCell.Worksheet.AutoFilter.Range.Column 'オートフィルターが適用される範囲の左端の列番号を取得

    YE = ActiveCell.Worksheet.AutoFilter.Range.Rows(ActiveCell.Worksheet.AutoFilter.Range.Rows.Count).Row
    'オートフィルターが適用される範囲の下端の行番号を取得

    If YP < YS Or YP > YE Then
        MsgBox "オートフィルター範囲外（上・下方向）には貼り付けできません！"
        Exit Sub
    End If
    
    XP = XP - XS
    YP = YP - YS
    
    With ActiveCell.Worksheet.AutoFilter.Range
        Set rngDest = .Columns(XP)
        Set rngDest = Intersect(rngDest, rngDest.Offset(YP))
        If YP + YS = YE Then    '現在選択しているセルがオートフィルターが適用される範囲の下端の場合は可視セルの取得処理をパスする
        Else
            Set rngDest = rngDest.SpecialCells(xlCellTypeVisible)   '最後に可視セルを取得するのがみそ
        End If
    End With
    
    Set Dobj = New DataObject
    With Dobj
        .GetFromClipboard
        On Error Resume Next
        V = .GetText
        On Error GoTo 0
    End With
    
    If Not IsEmpty(V) Then    'クリップボードからテキストが取得できた時のみ実行
        V = Split(CStr(V), vbCrLf)
        i = 0
        For Each R In rngDest.Cells
            If V(i) = "" Then   '空白行がある場合のエラーを回避
                A = CStr(V(i))
                R.Value = A
            Else
                A = Split(CStr(V(i)), vbTab)
                R.Resize(, UBound(A) + 1).Value = A
            End If
            
            i = i + 1
            If i > UBound(V) Then Exit For
        Next
    Else
        MsgBox "クリップボードにデータがありません！"
        Exit Sub
    End If
    
    Set Dobj = Nothing
    Set rngDest = Nothing
    Set R = Nothing
End Sub