Sub CBから特巡へ電気所名を転記()
    'セル内を部分的にコピーすることを想定しています
    'コピー後に実行してください
    'ローカル的なマクロです
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    Dim 区切り As String
    Dim 全文字列 As String
    Dim 分割文字列 As Variant
    
    区切り = InputBox("区切り文字を指定して下さい。（セル内改行にする場合は空白でＯＫ又はキャンセル）")
    If 区切り = "" Then 区切り = vbLf
    
    myLib.GetFromClipboard
    
    On Error Resume Next
    
    全文字列 = myLib.GetText
    
    On Error GoTo 0
    
    If 全文字列 <> "" Then
        分割文字列 = Split(全文字列, 区切り)
        Sheets("特巡").Select
        Range("D10:H10").ClearContents
        Range("D10").Resize(1, UBound(分割文字列) + 1).Value = 分割文字列

    Else
        MsgBox "クリップボードにデータがありません！"

    End If

End Sub