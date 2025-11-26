Sub カタカナのみ半角にする()
    '選択範囲に対して処理を行います
    '「－」は、ひらがなとカタカナ区別なく半角にします。半角にしたくない場合、ァ-ヶの後のーを消去して下さい
    Dim myRange As Range
    Dim i As Integer

    For Each myRange In Selection
    
        i = 1
    
        Do While i <= Len(myRange.Value)
            If Mid(myRange.Value, i, 1) Like "[ァ-ヶー]" Then
                myRange.Value = Application.WorksheetFunction.Replace(myRange.Value, i, 1, _
                StrConv(Mid(myRange.Value, i, 1), vbNarrow))
            End If
            
            i = i + 1
            
        Loop
    Next myRange
    
End Sub