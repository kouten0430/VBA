Sub 選択中の段落の先頭にクリップボードのデータを貼り付け()
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    Dim 全文字列 As String
    Dim 分割文字列 As Variant
    Dim i As Integer
    Dim 段落 As Paragraph
    
    myLib.GetFromClipboard
    
    On Error Resume Next
    
    全文字列 = myLib.GetText
    
    On Error GoTo 0
    
    If 全文字列 <> "" Then
        分割文字列 = Split(全文字列, vbCrLf)  '全文字列を改行で分割し配列に格納する
        i = 0
    
        For Each 段落 In Selection.Paragraphs
            If i <= UBound(分割文字列) Then
                段落.Range.InsertBefore 分割文字列(i)
                i = i + 1
            Else
                Exit For
            End If
        Next
        
    Else
        MsgBox "クリップボードにデータがありません！"

    End If
    
End Sub