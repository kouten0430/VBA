Sub 前のインデックスのブックをアクティブにする()
    '個人用マクロブックはインデックス1
    '2以降はブックを開いた順
    Dim i As Integer
    
    For i = Workbooks.Count To 3 Step -1    '専用のプロパティがないため、ループでアクティブブックのインデックスを調べる
        If Workbooks(i).Name = ActiveWorkbook.Name Then
            Workbooks(i - 1).Activate
            Exit For
        End If
    Next i
    
End Sub