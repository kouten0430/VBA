Sub もう片方の文書をアクティブにする()
    Dim i As Integer
    
    If Documents.Count = 2 Then
        For i = 1 To Documents.Count    '専用のプロパティがないため、ループでアクティブな文書のインデックスを調べる
            If Documents(i).Name = ActiveDocument.Name Then
                If i = 1 Then
                    Documents(i + 1).Activate
                    Exit For
                    
                ElseIf i = 2 Then
                    Documents(i - 1).Activate
                    Exit For
    
                End If
                
            End If
            
        Next i
        
    Else
        MsgBox "文書を2つだけ開いた状態にして下さい。"
        
    End If
    
End Sub