Sub 非連続な位置にクリップボードのデータを貼り付け()
    'クリップボードのデータを貼り付ける位置にあらかじめ目印をつけておいて下さい
    '目印１個がクリップボードデータの１行分に対応します
    Dim Mejirushi As String
    Dim 全文字列 As String
    Dim 分割文字列 As Variant
    Dim i As Integer
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    Mejirushi = "(＠_＠;)"  '検索する文字列（目印）
    
    myLib.GetFromClipboard
    
    On Error Resume Next
    
    全文字列 = myLib.GetText
    
    On Error GoTo 0
    
    If 全文字列 <> "" Then
        分割文字列 = Split(全文字列, vbCrLf)  '全文字列を改行で分割し、配列に格納する
        i = 0
        
        ActiveDocument.Range(0, 0).Select   '文書の先頭から検索を開始する
    
        With Selection.Find
            .Text = Mejirushi
            
            Do While .Execute   '検索に一致する文字列が無くなるまで下方向に検索する
                If i <= UBound(分割文字列) Then
                    Selection.Range.Text = 分割文字列(i)
                    i = i + 1
                Else
                    Selection.Range.Text = ""   '検索の途中で配列の中身が無くなった場合、余った目印は空白に置換する
                End If
            Loop

        End With
    
    Else
        MsgBox "クリップボードにデータがありません！"

    End If
    
End Sub