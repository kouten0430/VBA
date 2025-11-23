Sub オートフィルターで絞り込みされている列を表示する()
    Dim i As Integer
    Dim tmp As Integer
    
    If ActiveSheet.FilterMode = True Then
        For i = 1 To ActiveSheet.AutoFilter.Filters.Count
            If ActiveSheet.AutoFilter.Filters(i).On = True Then
                ActiveWindow.ScrollColumn = ActiveSheet.AutoFilter.Range.Column + i - 1
                
                tmp = MsgBox("この列でよければ「はい」を、" & vbCrLf & "次の列を表示するには「いいえ」を押して下さい。", vbYesNo)
                
                If tmp = vbYes Then  '「はい」を選択した場合はプロシージャを終了する
                    Exit Sub
                End If
                
            End If
            
        Next i
        
        MsgBox "次の列はありません。"
        
    Else
        MsgBox "絞り込みされている箇所はありません。"
        
    End If
    
End Sub