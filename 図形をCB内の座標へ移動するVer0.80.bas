Sub 図形をCB内の座標へ移動する()
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    Dim CB As String
    Dim 配列 As Variant
    Dim tmp As String

    myLib.GetFromClipboard
    
    On Error Resume Next
    
    CB = myLib.GetText
    
    On Error GoTo 0
    
    If CB <> "" Then
        配列 = Split(CB, vbCrLf)

    Else
        MsgBox "クリップボードにデータがありません！"
        Exit Sub

    End If

    If Selection.ShapeRange.Count = 1 Then
        tmp = InputBox("どのように移動するか選んで下さい。" & vbCrLf & vbCrLf & _
        "1:水平方向：左端合わせ" & vbCrLf & _
        "2:水平方向：中心合わせ" & vbCrLf & _
        "3:水平方向：右端合わせ" & vbCrLf & _
        "4:垂直方向：下端合わせ" & vbCrLf & _
        "5:垂直方向：中心合わせ" & vbCrLf & _
        "6:垂直方向：上端合わせ" & vbCrLf & _
        "7:中心合わせ")
        
        If tmp = "" Then
            Exit Sub
        ElseIf tmp = "1" Then
            Selection.ShapeRange.Left = 配列(0)
        ElseIf tmp = "2" Then
            Selection.ShapeRange.Left = 配列(0) - Selection.ShapeRange.Width / 2
        ElseIf tmp = "3" Then
            Selection.ShapeRange.Left = 配列(0) - Selection.ShapeRange.Width
        ElseIf tmp = "4" Then
            Selection.ShapeRange.Top = 配列(1) - Selection.ShapeRange.Height
        ElseIf tmp = "5" Then
            Selection.ShapeRange.Top = 配列(1) - Selection.ShapeRange.Height / 2
        ElseIf tmp = "6" Then
            Selection.ShapeRange.Top = 配列(1)
        ElseIf tmp = "7" Then
            Selection.ShapeRange.Left = 配列(0) - Selection.ShapeRange.Width / 2
            Selection.ShapeRange.Top = 配列(1) - Selection.ShapeRange.Height / 2
        Else
            MsgBox "指定された数字で入力して下さい。"
            Exit Sub
        End If
        
    Else
        MsgBox "図形が複数選択されています。"
        
    End If

End Sub