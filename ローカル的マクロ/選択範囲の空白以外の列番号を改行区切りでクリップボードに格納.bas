Sub 選択範囲の空白以外の列番号を改行区切りでクリップボードに格納()
    '「CB内列番号のいずれかと一致するセルをClearContentsする」と合わせて使用します
    'ローカル的なマクロです
    Dim myRange As Range
    Dim V As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    If Selection.Count > 1 Then
        For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
            If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上の値のみ取り出す
                If myRange.Value <> "" Then V = V & myRange.Column & vbCrLf
            End If
        Next myRange
        
        V = Left(V, Len(V) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）
        
    Else
        V = ActiveCell.Column
        
    End If
    
    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する

End Sub