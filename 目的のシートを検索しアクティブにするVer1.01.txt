Sub 目的のシートを検索しアクティブにする()
    '大文字と小文字を区別しません｡
    '半角と全角を区別しません｡
    Dim sw As String
    Dim myWs As Worksheet
    Dim mb As Integer
    
    sw = InputBox("検索する文字列")
        If sw = "" Then '空白でOKした場合、キャンセルした場合の処理
            Exit Sub
        End If
    
    For Each myWs In Worksheets
        If StrConv(StrConv(myWs.Name, vbLowerCase), vbNarrow) Like "*" & _
        StrConv(StrConv(sw, vbLowerCase), vbNarrow) & "*" Then   '半角小文字に統一し、部分一致条件でシート名を検索する
            myWs.Activate
            
            mb = MsgBox(prompt:="このシートでよければ「はい」を、" & vbCrLf & _
            "次を検索するには「いいえ」を押して下さい。", Buttons:=vbYesNo + vbDefaultButton2)
            
            If mb = 6 Then  '「はい」を選択した場合はプロシージャを終了する
                Exit Sub
            End If
            
        End If
    Next myWs
    
    MsgBox "検索に一致するものはありませんでした。"
    
End Sub