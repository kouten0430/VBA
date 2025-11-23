Sub 設備停止CSVから月間予定へ転記()
    '前段として設備停止CSVから新変保守が対応するものを抽出するを実行しておく
    '転記元と転記先のブックを開いておく
    '転記先の該当月のシートを選択しておく
    '転記元のブックをアクティブにし、カレントリージョンが効く位置でマクロを実行する
    'Cellsメソッドの失敗を防ぐため、転記元のCellsはオブジェクトを省略せずに明示してます
    'ローカル的なマクロです
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 列始 As Long
    Dim 列終 As Long
    Dim i As Long
    Dim j As Long
    Dim 開始日時の列番号 As Long
    Dim 終了日時の列番号 As Long
    Dim 実施箇所の列番号 As Long
    Dim 作業内容の列番号 As Long
    Dim 要求箇所の列番号 As Long
    Dim 毎連区分の列番号 As Long
    Dim 開始日 As Date
    Dim 終了日 As Date
    Dim 開始日時 As String
    Dim 終了日時 As String
    Dim 実施箇所 As String
    Dim 作業内容 As String
    Dim 要求箇所 As String
    Dim 毎連区分 As String
    Dim 一致 As Range
    Dim 転記元ワークブック名 As String
    Dim 転記先ワークブック名 As String
    Dim 月行 As Long
    Dim 月初 As Long
    Dim 月末 As Long
    Dim 正式名列 As Long
    
    転記元ワークブック名 = ActiveWorkbook.Name
    転記先ワークブック名 = "a.xlsx"
    月行 = 3    '転記先の日付が入っている行番号
    月初 = 5    '転記先の月初の日付が入っている列番号
    月末 = 35   '転記先の月末の日付が入っている列番号
    正式名列 = 4    '転記先の正式な電気所名が入っている列番号

    '---ここから行番号・列番号を取得する処理---
    
    行始 = ActiveCell.CurrentRegion.Row
    行終 = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row
    列始 = ActiveCell.CurrentRegion.Column
    列終 = ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column
    
    For i = 列始 To 列終
        If Cells(行始, i).Value = "要求期間　開始日時" Then
            開始日時の列番号 = i
        ElseIf Cells(行始, i).Value = "要求期間　終了日時" Then
            終了日時の列番号 = i
        ElseIf Cells(行始, i).Value = "実施箇所" Then
            実施箇所の列番号 = i
        ElseIf Cells(行始, i).Value = "作業内容" Then
            作業内容の列番号 = i
        ElseIf Cells(行始, i).Value = "要求箇所" Then
            要求箇所の列番号 = i
        ElseIf Cells(行始, i).Value = "要求期間　毎連区分" Then
            毎連区分の列番号 = i
        End If
    Next i
    
    '---ここから該当電気所のデータを転記する処理---
    
    Application.ScreenUpdating = False
    
    For i = 行始 + 1 To 行終
        If Workbooks(転記元ワークブック名).Sheets(1).Cells(i, "A").Interior.Color = 15921906 Then
            開始日 = WorksheetFunction.Floor(Workbooks(転記元ワークブック名).Sheets(1).Cells(i, 開始日時の列番号).Value, 1)   '時分は切り捨てる
            終了日 = WorksheetFunction.Floor(Workbooks(転記元ワークブック名).Sheets(1).Cells(i, 終了日時の列番号).Value, 1)   '時分は切り捨てる
            
            開始日時 = Format(Workbooks(転記元ワークブック名).Sheets(1).Cells(i, 開始日時の列番号).Value, "m/d h:mm")
            If 開始日 = 終了日 Then
                終了日時 = Format(Workbooks(転記元ワークブック名).Sheets(1).Cells(i, 終了日時の列番号).Value, "h:mm")
            Else
                終了日時 = Format(Workbooks(転記元ワークブック名).Sheets(1).Cells(i, 終了日時の列番号).Value, "m/d h:mm")
            End If
            実施箇所 = Workbooks(転記元ワークブック名).Sheets(1).Cells(i, 実施箇所の列番号).Value
            作業内容 = Workbooks(転記元ワークブック名).Sheets(1).Cells(i, 作業内容の列番号).Value
            要求箇所 = Workbooks(転記元ワークブック名).Sheets(1).Cells(i, 要求箇所の列番号).Value
            毎連区分 = Workbooks(転記元ワークブック名).Sheets(1).Cells(i, 毎連区分の列番号).Value
        
            Workbooks(転記先ワークブック名).Activate
            
            Set 一致 = Range(Cells(1, 正式名列), Cells(Cells(Rows.Count, 正式名列).End(xlUp).Row, 正式名列)).Find(実施箇所, LookAt:=xlWhole)
            
            If Not 一致 Is Nothing Then
                For j = 月初 To 月末
                    If Cells(月行, j).Value >= 開始日 And Cells(月行, j).Value <= 終了日 And TypeName(Cells(月行, j).Value) = "Date" Then
                        If Cells(一致.Row, j).Value <> "" Then 'セルが空白でなければ改行し、空白であれば改行しない
                            Cells(一致.Row, j).Value = Cells(一致.Row, j).Value & vbLf & "--------" & vbLf & "【" & 要求箇所 & "】" & 作業内容 & "　" & 開始日時 & " ～ " & 終了日時 & "（" & 毎連区分 & "）"
                        Else
                            Cells(一致.Row, j).Value = "【" & 要求箇所 & "】" & 作業内容 & "　" & 開始日時 & " ～ " & 終了日時 & "（" & 毎連区分 & "）"
                        End If
                        
                        Cells(一致.Row, j).Interior.Color = 15921906
                        
                        Workbooks(転記元ワークブック名).Sheets(1).Cells(i, 実施箇所の列番号).Interior.Color = 65535
    
                    End If
                Next j
                
            End If
            
        End If
    Next i
    
End Sub