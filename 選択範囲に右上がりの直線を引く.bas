Sub 選択範囲に右上がりの直線を引く()
    Dim x As Single
    Dim y As Single
    Dim xe As Single
    Dim ye As Single

    x = Selection.Left
    y = Selection.Top + Selection.Height
    xe = Selection.Left + Selection.Width
    ye = Selection.Top

    ActiveSheet.Shapes.AddLine x, y, xe, ye
    
End Sub