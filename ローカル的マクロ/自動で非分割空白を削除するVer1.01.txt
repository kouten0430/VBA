Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Cells(1, 1).Address <> "$A$5" Then Exit Sub

    Cells.Replace ChrW(160), "", xlPart
    
    AutoFilter.Sort.SortFields.Clear
    AutoFilter.Sort.SortFields.Add Key:=Range("J1"), Order:=xlDescending
    AutoFilter.Sort.Apply

End Sub