Sub 数式が入力されているセルのみを選択する()
    '選択範囲内で数式が入力されているセルのみを再選択します
    Dim myRange As Range
    Dim myUni As Range

    For Each myRange In Selection
        If myRange.HasFormula Then
            If myUni Is Nothing Then
                Set myUni = myRange    '検索に一致した一番最初のセル
            Else
                Set myUni = Union(myUni, myRange)    '検索に一致した二番目以降のセル
            End If
        End If
    Next myRange
    
    If Not myUni Is Nothing Then
        myUni.Select    '検索に一致したセルをまとめて選択する
    End If

End Sub