Sub カタカナのみ全角にする()
    '選択範囲に対して処理を行います
    Dim myRange As Range
    Dim i As Integer
    Dim cnt As Integer
    Dim 濁 As Integer

    For Each myRange In Selection
    
        i = 1
    
        Do While i <= Len(myRange.Value)
            If Mid(myRange.Value, i, 1) Like "[ｦ-ﾟ]" Then
                cnt = cnt + 1
                If Mid(myRange.Value, i, 1) Like "[ﾞﾟ]" Then
                    濁 = 濁 + 1 '濁点と半濁点の数をカウントする
                End If
            ElseIf cnt > 0 Then
                myRange.Value = Application.WorksheetFunction.Replace(myRange.Value, i - cnt, cnt, _
                StrConv(Mid(myRange.Value, i - cnt, cnt), vbWide))
                i = i - 濁: 濁 = 0 '濁点と半濁点の数をiから引く
                cnt = 0
            End If
            
            i = i + 1
            
        Loop
        
        If cnt > 0 Then
            myRange.Value = Application.WorksheetFunction.Replace(myRange.Value, i - cnt, cnt, _
            StrConv(Mid(myRange.Value, i - cnt, cnt), vbWide))
            濁 = 0
            cnt = 0
        End If

    Next myRange
    
End Sub