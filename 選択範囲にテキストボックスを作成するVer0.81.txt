Sub 選択範囲にテキストボックスを作成する()
    '選択範囲は複数指定可能です
    Dim i As Integer
    Dim x As Single
    Dim y As Single
    Dim xe As Single
    Dim ye As Single
    
    For i = 1 To Selection.Areas.Count
        x = Selection.Areas(i).Left
        y = Selection.Areas(i).Top
        xe = Selection.Areas(i).Width
        ye = Selection.Areas(i).Height
                
        ActiveSheet.Shapes.AddTextbox msoTextOrientationHorizontal, x, y, xe, ye

    Next i

End Sub