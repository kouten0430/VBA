Sub 水平方向に離れたセルのデータをCBに格納する列固定版列3()
    '選択したセルから水平方向に離れたセルのデータを改行区切りでCBに格納します
    '選択したセルにXをつけます
    '列のデータがある範囲を網かけします
    '選択した日の全てのデータに斜線が入っているかチェックします
    'ローカル的なマクロです
    Dim 列 As Long
    Dim 始 As Long
    Dim 終 As Long
    Dim myRange As Range
    Dim V As String
    Dim i As Long
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    列 = 3
    
    始 = Cells(1, Selection.Column).End(xlDown).Row
    終 = Cells(Rows.Count, Selection.Column).End(xlUp).Row
    
    For Each myRange In Selection
        If myRange.Value <> "" Then
            V = V & Cells(myRange.Row, 列).Value & vbCrLf
            
            myRange.Borders(xlDiagonalUp).LineStyle = True
            myRange.Borders(xlDiagonalDown).LineStyle = True
            
            Range(Cells(始, myRange.Column), Cells(終, myRange.Column)).Interior.Pattern = xlPatternGray8
            
        End If
        
    Next myRange
    
    If V <> "" Then
        V = Left(V, Len(V) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）
        
        myLib.SetText V  '変数の値をDataObjectに格納する
        myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する
        
    End If
    
    For i = Cells.Find("組織スケジュール", LookAt:=xlWhole).MergeArea(Cells.Find("組織スケジュール", LookAt:=xlWhole).MergeArea.Count).Row + 1 To Cells(Rows.Count, Cells.Find("組織スケジュール", LookAt:=xlWhole).Column).End(xlUp).Row
        If Cells(i, Selection.Column).Value <> "" And Cells(i, Selection.Column).Borders(xlDiagonalUp).LineStyle = xlLineStyleNone Then
            GoTo skip
            
        End If

    Next i
    
    MsgBox "全てのデータに斜線が入っています" & vbCrLf & "（又はデータ無し）"
    
skip:

End Sub