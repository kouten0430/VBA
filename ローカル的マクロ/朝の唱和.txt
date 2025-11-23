Sub 朝の唱和()
'ローカル的なマクロです
'技術事務の行と担当の列が交わるセルを選択して実行する
Dim V As String

V = "朝の唱和："    '必要に応じて変更する

Cells(2, ActiveCell.Column).Borders(xlDiagonalUp).LineStyle = True
Cells(2, ActiveCell.Column).Borders(xlDiagonalDown).LineStyle = True
Cells(ActiveCell.Row, "F").Value = Cells(ActiveCell.Row, "F").Value & vbCrLf & V & Cells(2, ActiveCell.Column).Value

End Sub