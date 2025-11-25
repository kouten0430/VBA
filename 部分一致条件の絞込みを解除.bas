Sub 部分一致条件の絞込みを解除()
    '絞り込みを行ったデータのアクティブセル領域内で実行
    '絞り込みを行った列で実行
    Rows(ActiveCell.CurrentRegion.Row & ":" & _
    ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row).Hidden = False
End Sub