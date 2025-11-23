Sub 実行中のタスク名をクリップボードに転送()
    'TasksコレクションはWordでのみ使用可能
    Dim myTask As Task
    Dim CB As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    For Each myTask In Tasks
        If myTask.Visible = True Then
            CB = CB & myTask.Name & vbCrLf
            
        End If
        
    Next myTask
    
    CB = Left(CB, Len(CB) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）
    
    myLib.SetText CB  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する

End Sub