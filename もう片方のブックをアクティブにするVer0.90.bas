Sub もう片方のブックをアクティブにする()
    '個人用マクロブックはインデックス1
    '2以降はブックを開いた順
    Dim i As Integer
    
    If Workbooks.Count = 3 Then
        For i = 1 To Workbooks.Count    '専用のプロパティがないため、ループでアクティブブックのインデックスを調べる
            If Workbooks(i).Name = ActiveWorkbook.Name Then
                If i = 2 Then
                    Workbooks(i + 1).Activate
                    Exit For
                    
                ElseIf i = 3 Then
                    Workbooks(i - 1).Activate
                    Exit For
    
                End If
                
            End If
            
        Next i
        
    Else
        MsgBox "ブックを2つだけ開いた状態にして下さい。"
        
    End If
    
End Sub