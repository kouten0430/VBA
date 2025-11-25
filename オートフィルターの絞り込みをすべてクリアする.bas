Sub オートフィルターの絞り込みをすべてクリアする()
    If ActiveSheet.FilterMode = True Then
        ActiveSheet.ShowAllData
        
    Else
        MsgBox "絞り込みされている箇所はありません。"
        
    End If
    
End Sub