Sub c•ûŒü‚ÉŒ‹‡()
    Dim i As Integer

    For i = Selection.Column To Selection.Columns(Selection.Columns.Count).Column
        Range(Cells(Selection.Row, i), Cells(Selection.Rows(Selection.Rows.Count).Row, i)).Merge
    Next i

End Sub