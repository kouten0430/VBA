Sub 配布予定表の開始日を設定する()
    'ローカル的なマクロです
    Dim 指定日 As Date
    Dim 列番号 As Long
    
    指定日 = DateValue(InputBox("曜日を取得する日を西暦/月/日で入力（例：2021/11/1）", , Year(Date) & "/" & Month(Date) + 1 & "/" & 1))
    列番号 = Cells.Find(Format(指定日, "aaa"), LookAt:=xlWhole).Column
    
    Range("B8").Value = 指定日 - (列番号 - 2)
    
End Sub