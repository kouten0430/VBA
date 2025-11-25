Sub セルの座標をCBに格納する()
    Dim tmp As String
    Dim x As Single
    Dim y As Single
    Dim CB As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する

    If Selection.Count = 1 Then
        tmp = InputBox("取得する位置を選んで下さい。" & vbCrLf & vbCrLf & _
        "1:左上" & vbCrLf & _
        "2:上中心" & vbCrLf & _
        "3:右上" & vbCrLf & _
        "4:左中心" & vbCrLf & _
        "5:右中心" & vbCrLf & _
        "6:左下" & vbCrLf & _
        "7:下中心" & vbCrLf & _
        "8:右下" & vbCrLf & _
        "9:中心")
        
        If tmp = "" Then
            Exit Sub
        ElseIf tmp = "1" Then
            x = Selection.Left
            y = Selection.Top
        ElseIf tmp = "2" Then
            x = Selection.Left + Selection.Width / 2
            y = Selection.Top
        ElseIf tmp = "3" Then
            x = Selection.Left + Selection.Width
            y = Selection.Top
        ElseIf tmp = "4" Then
            x = Selection.Left
            y = Selection.Top + Selection.Height / 2
        ElseIf tmp = "5" Then
            x = Selection.Left + Selection.Width
            y = Selection.Top + Selection.Height / 2
        ElseIf tmp = "6" Then
            x = Selection.Left
            y = Selection.Top + Selection.Height
        ElseIf tmp = "7" Then
            x = Selection.Left + Selection.Width / 2
            y = Selection.Top + Selection.Height
        ElseIf tmp = "8" Then
            x = Selection.Left + Selection.Width
            y = Selection.Top + Selection.Height
        ElseIf tmp = "9" Then
            x = Selection.Left + Selection.Width / 2
            y = Selection.Top + Selection.Height / 2
        Else
            MsgBox "指定された数字で入力して下さい。"
            Exit Sub
        End If
        
        CB = x & vbCrLf & y
        
        myLib.SetText CB  '変数の値をDataObjectに格納する
        myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する
        
    Else
        MsgBox "セルが複数選択されています。"
        
    End If

End Sub