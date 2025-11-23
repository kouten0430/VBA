Sub 結合セルの選択位置に関係なく左上の値のみ取り出す()
    'フィルタリングで結合セルの一部分しか表示されていない場合に使用する
    Dim myRange As Range
    Dim V As String
    Dim Addr As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する

    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
        If myRange.MergeArea(1).Address <> Addr Then    '結合セルは一度だけ処理を行う
            V = V & myRange.MergeArea(1).Value & vbCrLf '結合セルの選択位置に関係なく左上の値のみ取り出す
            Addr = myRange.MergeArea(1).Address
        End If
    Next myRange
    
    V = Left(V, Len(V) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）

    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する

End Sub