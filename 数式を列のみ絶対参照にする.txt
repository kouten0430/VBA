Sub ”®‚ğ—ñ‚Ì‚İâ‘ÎQÆ‚É‚·‚é()
    Dim myRange As Range

    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)
        If myRange.Value <> "" Then   'ƒZƒ‹‚Ì’l‚ª‹ó”’‚Ìê‡‚Íˆ—‚ğ‚µ‚È‚¢
            myRange.Formula = Application.ConvertFormula(Formula:=myRange.Formula, _
            FromReferenceStyle:=xlA1, ToReferenceStyle:=xlA1, ToAbsolute:=xlRelRowAbsColumn)
        End If
    Next myRange
    
End Sub