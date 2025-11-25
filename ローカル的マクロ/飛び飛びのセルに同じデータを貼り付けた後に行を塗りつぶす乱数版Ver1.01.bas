Sub 飛び飛びのセルに同じデータを貼り付けた後に行を塗りつぶす乱数版()
    'クリップボード内の同じデータを選択した飛び飛びのセルに貼り付けすることができます
    '結合セルを単一セルのように扱うことができます
    'ローカル的なマクロです
    Dim 列始 As Long
    Dim 列終 As Long
    Dim tmp As Integer
    Dim 色 As Long
    Dim V As Variant
    Dim myRange As Range
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    列始 = Range("B1").Column
    列終 = Range("FL1").Column
    
    tmp = WorksheetFunction.RandBetween(1, 7)
    
    If tmp = "1" Then
        色 = 16777164   '薄い青
    ElseIf tmp = "2" Then
        色 = 13434828   '薄い緑
    ElseIf tmp = "3" Then
        色 = 13434879   '薄い黄
    ElseIf tmp = "4" Then
        色 = 15654399   '薄いピンク
    ElseIf tmp = "5" Then
        色 = 16767468   '薄い紫
    ElseIf tmp = "6" Then
        色 = 15790320   '薄い灰
    ElseIf tmp = "7" Then
        色 = 14083324   '薄いオレンジ
    End If

    myLib.GetFromClipboard
        On Error Resume Next
    V = myLib.GetText
        On Error GoTo 0

    If Not IsEmpty(V) Then
        If Selection.Count > 1 Then
            For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
                If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルにのみ処理を行う
                    myRange = V
                    Range(Cells(myRange.Row, 列始), Cells(myRange.MergeArea(myRange.MergeArea.Count).Row, 列終)).Interior.Color = 色
                    
                End If
        
            Next myRange
            
        Else
            ActiveCell.Value = V
            Range(Cells(ActiveCell.Row, 列始), Cells(ActiveCell.Row, 列終)).Interior.Color = 色
            
        End If

    Else
        MsgBox "クリップボードにデータがありません！"

    End If
    
End Sub