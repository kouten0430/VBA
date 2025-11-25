Sub 三つ以上の完全一致条件で絞込みを行う()
    'クリップボードの文字列を参照し、OR条件で絞込みします
    '現在選択しているセルの列をフィルタリングします
    'シートにオートフィルターがない場合は、そのセルを含むアクティブセル領域をオートフィルターに設定した上で絞込みします
    Dim XS As Integer
    Dim XP As Integer
    Dim V As Variant
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
        ActiveCell.AutoFilter Field:=XP, Criteria1:=V, Operator:=xlFilterValues
    Else
        MsgBox "クリップボードにデータがありません！"
    End If

End Sub