Sub 選択範囲のデータをカンマ区切りでクリップボードに格納()
    'Microsoft Forms 2.0 Object Libraryを参照設定して下さい
    
    Dim myRange As Range
    Dim V As String

    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上の値のみ取り出す
            V = V & myRange.Value & ","
        End If
    Next myRange
    
    V = Left(V, Len(V) - 1) '最終行のカンマ区切りを取り除く
    
    With New MSForms.DataObject
        .SetText V  '変数の値をDataObjectに格納する
        .PutInClipboard 'DataObjectのデータをクリップボードに格納する
    End With

End Sub