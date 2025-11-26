Sub 選択した図形内の文字を置換する()
    Dim 置換前 As String
    Dim 置換後 As String
    Dim 図形 As Shape
    
    置換前 = InputBox("置換前の文字を入力")
    If 置換前 = "" Then Exit Sub
    
    置換後 = InputBox("置換後の文字を入力")
    If 置換後 = "" Then Exit Sub
    
    For Each 図形 In Selection.ShapeRange
        図形.TextFrame.Characters.Text = Replace(図形.TextFrame.Characters.Text, 置換前, 置換後)
        
    Next 図形
    
End Sub