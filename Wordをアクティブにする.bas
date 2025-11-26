Sub Wordをアクティブにする()
    On Error GoTo エラー処理
    
    AppActivate "- Word"
    
    On Error GoTo 0
    
エラー処理:

End Sub