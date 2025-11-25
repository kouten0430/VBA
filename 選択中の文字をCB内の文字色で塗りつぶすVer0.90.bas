Sub 選択中の文字をCB内の文字色で塗りつぶす()
    Dim i As Integer
    Dim 色 As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    
    myLib.GetFromClipboard
    
    On Error Resume Next
    
    色 = myLib.GetText
    
    On Error GoTo 0

    If 色 <> "" Then
        For i = 1 To Selection.Areas.Count
            Selection.Areas(i).Font.Color = 色
            
        Next i
        
    Else
        MsgBox "クリップボードにデータがありません！"
        
    End If
    
End Sub