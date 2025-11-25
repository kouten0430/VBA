Sub 指定文字で区切りクリップボードへ転送する()
    'クリップボードへ転送したいセルを選択（複数可）した状態で実行して下さい
    'セルを選択した順にデータを右方向へ連結していきます
    'InputBoxがブランクなら、区切り文字なしで連結していきます
    Dim myRange As Range
    Dim V As String
    Dim dc As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    dc = Application.InputBox(Prompt:="区切り文字を指定して下さい。" & vbCrLf & "（改行にする場合はキャンセルして下さい）", Type:=2)
    
        If dc = "False" Then
            dc = vbCrLf   'キャンセルの場合は区切りを改行にする
        End If

    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上の値のみ取り出す
            V = V & myRange.Value & dc
        End If
    Next myRange
    
    V = Left(V, Len(V) - Len(dc)) '最後の区切り文字を取り除く

    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する

End Sub