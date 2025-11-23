Sub セル色の値を改行区切りでクリップボードに格納()
    '格納する値はリトルエンディアンです
    
    Dim myRange As Range
    Dim V As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    If Selection.Count > 1 Then
        For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
            If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上の値のみ取り出す
                V = V & myRange.Interior.Color & vbCrLf
            End If
        Next myRange
        
        V = Left(V, Len(V) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）
        
    Else
        V = ActiveCell.Interior.Color
        
    End If
    
    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する

End Sub