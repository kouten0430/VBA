Sub 現在開いている全文書の同じ行にクリップボードのデータを貼り付ける()
    '現在アクティブになっている文書のカーソル位置と同じ行（厳密には段落）にデータが貼り付けされます
    'データが貼り付けされる順番は文書を開いた順番と逆（最後に開いた文書 → 最初に開いた文書）になります
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    Dim 全文字列 As String
    Dim 分割文字列 As Variant
    Dim i As Integer
    Dim 段落番号 As Integer
    Dim 文書 As Document
    
    myLib.GetFromClipboard
    
    On Error Resume Next
    
    全文字列 = myLib.GetText
    
    On Error GoTo 0
    
    If 全文字列 <> "" Then
        分割文字列 = Split(全文字列, vbCrLf)  '全文字列を改行で分割し配列に格納する
        i = 0
    
        段落番号 = ActiveDocument.Range(0, Selection.End + 1).Paragraphs.Count  'カーソル位置の段落番号を取得する
    
        For Each 文書 In Documents  '処理の順番は文書を開いた順番と逆になることに注意！
            If i <= UBound(分割文字列) Then
                文書.Paragraphs(段落番号).Range.InsertBefore 分割文字列(i)
                i = i + 1
            Else
                Exit For
            End If

        Next
        
    Else
        MsgBox "クリップボードにデータがありません！"

    End If
    
End Sub