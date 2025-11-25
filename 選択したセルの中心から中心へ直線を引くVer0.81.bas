Sub 選択したセルの中心から中心へ直線を引く()
    Dim i As Integer
    Dim x As Single
    Dim y As Single
    Dim xe As Single
    Dim ye As Single

    If Selection.Count = Selection.Areas.Count Then
        For i = 1 To Selection.Count
            If i > 1 Then
                x = Selection.Areas(i - 1).Left + Selection.Areas(i - 1).Width / 2
                y = Selection.Areas(i - 1).Top + Selection.Areas(i - 1).Height / 2
                xe = Selection.Areas(i).Left + Selection.Areas(i).Width / 2
                ye = Selection.Areas(i).Top + Selection.Areas(i).Height / 2
                
                ActiveSheet.Shapes.AddLine x, y, xe, ye
                
            End If
            
        Next i
        
    End If

End Sub