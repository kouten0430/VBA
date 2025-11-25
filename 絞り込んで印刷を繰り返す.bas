Sub 絞り込んで印刷を繰り返す()
    'クリップボードの文字列を配列に取り込み、配列の内容で順番に絞込みします
    '現在選択しているセルの列をフィルタリングします
    'シートにオートフィルターがない場合は、そのセルを含むアクティブセル領域をオートフィルターに設定した上で絞込みします
    '現在の印刷設定で印刷します
    Dim XS As Integer
    Dim XP As Integer
    Dim YS As Long
    Dim YE As Long
    Dim V As Variant
    Dim i As Integer
    Dim 可視セル数 As Long
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    myLib.GetFromClipboard
        On Error Resume Next
    V = myLib.GetText
        On Error GoTo 0
    
    If Not IsEmpty(V) Then
        V = Split(CStr(V), vbCrLf)
        ActiveCell.AutoFilter Field:=1  '引数は既にオートフィルターがある場合に解除しないためのダミー
        XP = ActiveCell.Column  '現在選択しているセルの列番号を取得
        XS = ActiveCell.Worksheet.AutoFilter.Range.Column 'オートフィルターが適用される範囲の左端の列番号を取得
        XP = XP + 1 - XS    '抽出条件の対象となる列番号
        YS = ActiveCell.Worksheet.AutoFilter.Range.Row 'オートフィルターが適用される範囲の上端の行番号を取得
        YE = ActiveCell.Worksheet.AutoFilter.Range.Rows(ActiveCell.Worksheet.AutoFilter.Range.Rows.Count).Row   'オートフィルターが適用される範囲の下端の行番号を取得
        
        i = 0
        
        Do While i <= UBound(V)
            ActiveCell.AutoFilter Field:=XP, Criteria1:=V(i), Operator:=xlFilterValues
            可視セル数 = Range(Cells(YS, XP), Cells(YE, XP)).Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count
            
            If 可視セル数 > 1 Then ActiveSheet.PrintOut: DoEvents '絞り込みに一致するものがあった場合のみ印刷する
            
            i = i + 1
        Loop
    Else
        MsgBox "クリップボードにデータがありません！"
    End If

End Sub