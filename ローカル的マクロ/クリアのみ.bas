Sub クリアのみ()
    '選択中のセルの行に対して処理を行います
    'ローカル的なマクロです
    Range(Cells(Selection.Row, 3), Cells(Selection.Row, 6)).ClearContents
    Range(Cells(Selection.Row, 8), Cells(Selection.Row, 57)).ClearContents
End Sub