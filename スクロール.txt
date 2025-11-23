Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim rng
    If Target.Address <> "$B$1" Then Exit Sub
    Set rng = Range("B2:B65536").Find(Target.Value)
    ActiveWindow.ScrollRow = rng.Row
End Sub