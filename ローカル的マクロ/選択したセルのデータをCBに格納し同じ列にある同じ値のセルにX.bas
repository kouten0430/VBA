Sub 選択したセルのデータをCBに格納し同じ列にある同じ値のセルにX()
    '選択したセルのデータを改行区切りでCBに格納します
    '選択したセルと同じ列にある同じ値のセルにXをつけます
    '列のデータがある範囲を網かけします
    'ローカル的なマクロです
    Dim 始 As Long
    Dim 終 As Long
    Dim i As Long
    Dim 日付行 As Long
    Dim myRange As Range
    Dim V As String
    Dim j As Long
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    '---ここから行番号・列番号を取得する処理---
    
    始 = Cells(1, Selection.Column).End(xlDown).Row
    終 = Cells(Rows.Count, Selection.Column).End(xlUp).Row
    
    For i = 1 To Cells.Find("組織スケジュール", LookAt:=xlWhole).Row
        If TypeName(Cells(i, "S").Value) = "Date" Then
            日付行 = i
            Exit For
        End If
        
    Next i
    
    '---ここからする処理---
    
    For Each myRange In Selection
        V = V & myRange.Value & vbCrLf
        
        For j = 日付行 + 2 To Cells.Find("組織スケジュール", LookAt:=xlWhole).Row - 1
            If Cells(j, myRange.Column).Value = myRange.Value Then
                Cells(j, myRange.Column).Borders(xlDiagonalUp).LineStyle = True
                Cells(j, myRange.Column).Borders(xlDiagonalDown).LineStyle = True
                
            End If
            
        Next j
        
        Range(Cells(始, myRange.Column), Cells(終, myRange.Column)).Interior.Pattern = xlPatternGray8
        
    Next myRange
        
    V = Left(V, Len(V) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）
    
    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する

End Sub