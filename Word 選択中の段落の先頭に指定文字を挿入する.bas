Sub 選択中の段落の先頭に指定文字を挿入する()
    Dim 指定文字 As String
    Dim 段落 As Paragraph
    
    指定文字 = InputBox("先頭に挿入する文字を入力して下さい。")
        If 指定文字 = "" Then
            Exit Sub
        End If
    
    For Each 段落 In Selection.Paragraphs
        段落.Range.InsertBefore 指定文字
    Next
    
End Sub