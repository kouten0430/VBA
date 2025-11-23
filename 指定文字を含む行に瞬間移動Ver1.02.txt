Sub 指定文字を含む行に瞬間移動()
    '現在の表示位置から検索を開始します
    '行方向にのみ移動します
    '同じ行に複数一致セルがある場合は次の行まで処理をパスします
    Dim 指定文字 As String
    Dim 一致セル As Range
    Dim 最初の一致セル As Range
    Dim 前検索結果の行 As Long
    Dim tmp As Integer
    
    指定文字 = InputBox("指定文字を入力（部分一致）")
    If 指定文字 = "" Then Exit Sub
    
    Set 一致セル = Cells.Find(指定文字, After:=ActiveWindow.VisibleRange.Item(1, 1), LookAt:=xlPart, SearchOrder:=xlByRows)
    
    If 一致セル Is Nothing Then
        MsgBox "検索に一致するものはありませんでした。"
        Exit Sub
        
    End If
    
    Set 最初の一致セル = 一致セル
    
    Do
        If 前検索結果の行 <> 一致セル.Row Then
            ActiveWindow.ScrollRow = 一致セル.Row
            
            前検索結果の行 = 一致セル.Row
            
            tmp = MsgBox("この行でよければ「はい」を、" & vbCrLf & "次の行を表示するには「いいえ」を押して下さい。", vbYesNo + vbDefaultButton2)
            
            If tmp = vbYes Then  '「はい」を選択した場合はプロシージャを終了する
                Exit Do
            End If
            
            Set 一致セル = Cells.FindNext(一致セル)
            
            If 一致セル.Address = 最初の一致セル.Address Then
                MsgBox "すべて検索し終えました。"
                Exit Do
                
            End If
            
        Else
            Set 一致セル = Cells.FindNext(一致セル)
            
            If 一致セル.Address = 最初の一致セル.Address Then
                MsgBox "すべて検索し終えました。"
                Exit Do
                
            End If
        
        End If

    Loop
    
End Sub