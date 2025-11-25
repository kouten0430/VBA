Sub 設備停止CSVから入力シートへ転記()
    '転記元と転記先のブックを開いておく
    '転記元のブックをアクティブにし、転記したい行を選択してマクロを実行する
    'オートフィルタで何らかの絞り込みがされている時に非表示となっている列は、Findによる検索ができないので注意
    Dim i As Integer
    Dim 転記先シート As Worksheet
    Dim 列停区 As Long
    Dim 列作内 As Long
    Dim 列要開 As Long
    Dim 列要終 As Long
    Dim 列復設1 As Long
    Dim 列復時1 As Long
    Dim 列復設2 As Long
    Dim 列復時2 As Long
    Dim 列復設3 As Long
    Dim 列復時3 As Long
    Dim 列給連氏名 As Long
    Dim 列給連TEL As Long
    Dim 列要注 As Long
    
    For i = 1 To Workbooks.Count    '専用のプロパティがないため、ループでアクティブブックのインデックスを調べる
        If Workbooks(i).Name Like "新・設備停止作業計画支援システム入力シート*" Then
            Set 転記先シート = Workbooks(i).ActiveSheet
            Exit For
            
        End If
        
    Next i

    列停区 = Cells.Find("停止／充電区間", LookAt:=xlWhole).Column
    列作内 = Cells.Find("作業内容", LookAt:=xlWhole).Column
    列要開 = Cells.Find("要求期間　開始日時", LookAt:=xlWhole).Column
    列要終 = Cells.Find("要求期間　終了日時", LookAt:=xlWhole).Column
    列復設1 = Cells.Find("緊急復旧　復旧設備１", LookAt:=xlWhole).Column
    列復時1 = Cells.Find("緊急復旧　設備１復旧時間", LookAt:=xlWhole).Column
    列復設2 = Cells.Find("緊急復旧　復旧設備２", LookAt:=xlWhole).Column
    列復時2 = Cells.Find("緊急復旧　設備２復旧時間", LookAt:=xlWhole).Column
    列復設3 = Cells.Find("緊急復旧　復旧設備３", LookAt:=xlWhole).Column
    列復時3 = Cells.Find("緊急復旧　設備３復旧時間", LookAt:=xlWhole).Column
    列給連氏名 = Cells.Find("給電連絡責任者　氏名", LookAt:=xlWhole).Column
    列給連TEL = Cells.Find("給電連絡責任者　ＴＥＬ", LookAt:=xlWhole).Column
    列要注 = Cells.Find("要求時注釈", LookAt:=xlWhole).Column
    
    転記先シート.Range("A1").Value = Cells(ActiveCell.Row, 列停区).Value
    転記先シート.Range("A2").Value = Cells(ActiveCell.Row, 列作内).Value
    転記先シート.Range("A3").Value = Year(Cells(ActiveCell.Row, 列要開).Value)
    転記先シート.Range("A4").Value = Month(Cells(ActiveCell.Row, 列要開).Value)
    転記先シート.Range("A5").Value = Day(Cells(ActiveCell.Row, 列要開).Value)
    転記先シート.Range("A6").Value = Hour(Cells(ActiveCell.Row, 列要開).Value)
    転記先シート.Range("A7").Value = Minute(Cells(ActiveCell.Row, 列要開).Value)
    転記先シート.Range("A8").Value = Year(Cells(ActiveCell.Row, 列要終).Value)
    転記先シート.Range("A9").Value = Month(Cells(ActiveCell.Row, 列要終).Value)
    転記先シート.Range("A10").Value = Day(Cells(ActiveCell.Row, 列要終).Value)
    転記先シート.Range("A11").Value = Hour(Cells(ActiveCell.Row, 列要終).Value)
    転記先シート.Range("A12").Value = Minute(Cells(ActiveCell.Row, 列要終).Value)
    転記先シート.Range("A13").Value = Cells(ActiveCell.Row, 列復設1).Value
    転記先シート.Range("A14").Value = Cells(ActiveCell.Row, 列復時1).Value
    転記先シート.Range("A15").Value = Cells(ActiveCell.Row, 列復設2).Value
    転記先シート.Range("A16").Value = Cells(ActiveCell.Row, 列復時2).Value
    転記先シート.Range("A17").Value = Cells(ActiveCell.Row, 列復設3).Value
    転記先シート.Range("A18").Value = Cells(ActiveCell.Row, 列復時3).Value
    転記先シート.Range("A19").Value = Cells(ActiveCell.Row, 列給連氏名).Value
    転記先シート.Range("A20").Value = Cells(ActiveCell.Row, 列給連TEL).Value
    転記先シート.Range("A21").Value = Cells(ActiveCell.Row, 列要注).Value

End Sub