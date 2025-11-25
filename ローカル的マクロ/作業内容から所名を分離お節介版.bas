Sub 作業内容から所名を分離お節介版()
    '所名と作業内容が全角スペースで区切られていることが前提
    '作業内容の列で実行する
    'ローカル的なマクロです
    Dim myRange As Range
    Dim 配列 As Variant

    For Each myRange In Selection
        配列 = Split(myRange.Value, "　", 2)
        
        If 配列(0) = "岩瀬中央" Or 配列(0) = "富山火力" Or 配列(0) = "上滝" Then
            myRange.Offset(0, -4).Value = 配列(0) & "開閉所"
        Else
            myRange.Offset(0, -4).Value = 配列(0) & "変電所"
        End If
        
        myRange.Value = 配列(1)
        
    Next myRange
    
End Sub