Sub 選択中の段落のインデントを取得する()
    Dim s As Long
    Dim 始 As Integer
    
    s = Selection.Start + 1
    
    始 = ActiveDocument.Range(0, s).Paragraphs.Count

    MsgBox _
    "LeftIndent : " & ActiveDocument.Paragraphs(始).Range.ParagraphFormat.LeftIndent & vbCrLf & _
    "FirstLineIndent : " & ActiveDocument.Paragraphs(始).Range.ParagraphFormat.FirstLineIndent

End Sub