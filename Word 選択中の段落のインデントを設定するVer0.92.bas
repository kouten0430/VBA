Sub 選択中の段落のインデントを設定する()
    Dim 左 As String
    Dim 壱 As String
    Dim s As Long
    Dim e As Long
    Dim 始 As Integer
    Dim 終 As Integer
    Dim i As Integer
    
    左 = InputBox("LeftIndentを入力")
    壱 = InputBox("FirstLineIndentを入力")
    
    s = Selection.Start + 1
    e = Selection.End
    If s > e Then e = s
    
    始 = ActiveDocument.Range(0, s).Paragraphs.Count
    終 = ActiveDocument.Range(0, e).Paragraphs.Count
    
    For i = 始 To 終
        If 左 = "" Then 左 = ActiveDocument.Paragraphs(i).Range.ParagraphFormat.LeftIndent
        If 壱 = "" Then 壱 = ActiveDocument.Paragraphs(i).Range.ParagraphFormat.FirstLineIndent
        
        ActiveDocument.Paragraphs(i).Range.ParagraphFormat.CharacterUnitLeftIndent = 0  'CharacterUnitの値が優先されるためリセット
        ActiveDocument.Paragraphs(i).Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0 'CharacterUnitの値が優先されるためリセット
        
        ActiveDocument.Paragraphs(i).Range.ParagraphFormat.LeftIndent = 左
        ActiveDocument.Paragraphs(i).Range.ParagraphFormat.FirstLineIndent = 壱
        
    Next i

End Sub