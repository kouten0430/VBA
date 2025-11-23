Sub 数式を行のみ絶対参照にする()
    Dim myRange As Range

    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)
        If myRange.Value <> "" Then   'セルの値が空白の場合は処理をしない
            myRange.Formula = Application.ConvertFormula(Formula:=myRange.Formula, _
            FromReferenceStyle:=xlA1, ToReferenceStyle:=xlA1, ToAbsolute:=xlAbsRowRelColumn)
        End If
    Next myRange
    
End Sub