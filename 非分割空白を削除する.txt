Sub 非分割空白を削除する()
    Cells.Replace ChrW(160), "", xlPart
    
End Sub