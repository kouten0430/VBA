Sub Excelをアクティブにする()
    On Error GoTo エラー処理
    
    AppActivate "- Excel"
    
    On Error GoTo 0
    
エラー処理:

End Sub