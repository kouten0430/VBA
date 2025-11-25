Sub 作業内容から所名を分離プレーン版()
    '所名と作業内容が全角スペースで区切られていることが前提
    '作業内容の列で実行する
    'ローカル的なマクロです
    Dim myRange As Range
    Dim 配列 As Variant

    For Each myRange In Selection
        配列 = Split(myRange.Value, "　", 2)
        myRange.Offset(0, -4).Value = 配列(0)
        myRange.Value = 配列(1)
        
    Next myRange
    
End Sub