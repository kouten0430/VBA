Sub CB内の指定文字を制御文字に置換する()
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    Dim CB As String
    Dim 指定文字 As String
    Dim tmp As String
    Dim 制御文字 As String
    
    myLib.GetFromClipboard
    
    On Error Resume Next
    
    CB = myLib.GetText
    
    On Error GoTo 0
    
    If CB <> "" Then
        指定文字 = InputBox("指定文字を入力")
        If 指定文字 = "" Then Exit Sub
        
        tmp = InputBox("制御文字を選択" & vbCrLf & "1:CrLf" & vbCrLf & "2:Tab")
        If tmp = "" Then
            Exit Sub
        ElseIf tmp = "1" Then
            制御文字 = vbCrLf
        ElseIf tmp = "2" Then
            制御文字 = vbTab
        Else
            MsgBox "指定された数字で入力して下さい。"
            Exit Sub
        End If
        
        CB = Replace(CB, 指定文字, 制御文字)
        
        myLib.SetText CB  '変数の値をDataObjectに格納する
        myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する
        
    Else
        MsgBox "クリップボードにデータがありません！"

    End If
    
End Sub