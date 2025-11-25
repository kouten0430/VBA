Sub 選択中の段落の末尾に指定文字を挿入する()
    Dim 指定文字 As String
    Dim 段落 As Paragraph
    
    指定文字 = InputBox("末尾に挿入する文字を入力して下さい。")
        If 指定文字 = "" Then
            Exit Sub
        End If
    
    For Each 段落 In Selection.Paragraphs
       ActiveDocument.Range(0, 段落.Range.End - 1).InsertAfter 指定文字 '改行の手前に挿入するのがポイント
    Next
    
End Sub