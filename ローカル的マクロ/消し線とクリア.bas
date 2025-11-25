Sub 消し線とクリア()
    '選択中のセルの行に対して処理を行います
    'ローカル的なマクロです
    Range(Cells(Selection.Row, 3), Cells(Selection.Row, 11)).Font.Strikethrough = True
    Range(Cells(Selection.Row, 14), Cells(Selection.Row, 57)).ClearContents
End Sub