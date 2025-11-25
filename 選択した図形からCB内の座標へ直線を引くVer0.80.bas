Sub 選択した図形からCB内の座標へ直線を引く()
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    Dim CB As String
    Dim 配列 As Variant
    Dim tmp As String
    Dim 図形 As Shape
    Dim x As Single
    Dim y As Single
    Dim xe As Single
    Dim ye As Single
    
    myLib.GetFromClipboard
    
    On Error Resume Next
    
    CB = myLib.GetText
    
    On Error GoTo 0
    
    If CB <> "" Then
        配列 = Split(CB, vbCrLf)
        
        tmp = InputBox("図形のどの位置を始点とするか選んで下さい。" & vbCrLf & vbCrLf & _
        "1:左上" & vbCrLf & _
        "2:上中心" & vbCrLf & _
        "3:右上" & vbCrLf & _
        "4:左中心" & vbCrLf & _
        "5:右中心" & vbCrLf & _
        "6:左下" & vbCrLf & _
        "7:下中心" & vbCrLf & _
        "8:右下" & vbCrLf & _
        "9:中心")
        
        For Each 図形 In Selection.ShapeRange
            If tmp = "" Then
                Exit Sub
            ElseIf tmp = "1" Then
                x = 図形.Left
                y = 図形.Top
            ElseIf tmp = "2" Then
                x = 図形.Left + 図形.Width / 2
                y = 図形.Top
            ElseIf tmp = "3" Then
                x = 図形.Left + 図形.Width
                y = 図形.Top
            ElseIf tmp = "4" Then
                x = 図形.Left
                y = 図形.Top + 図形.Height / 2
            ElseIf tmp = "5" Then
                x = 図形.Left + 図形.Width
                y = 図形.Top + 図形.Height / 2
            ElseIf tmp = "6" Then
                x = 図形.Left
                y = 図形.Top + 図形.Height
            ElseIf tmp = "7" Then
                x = 図形.Left + 図形.Width / 2
                y = 図形.Top + 図形.Height
            ElseIf tmp = "8" Then
                x = 図形.Left + 図形.Width
                y = 図形.Top + 図形.Height
            ElseIf tmp = "9" Then
                x = 図形.Left + 図形.Width / 2
                y = 図形.Top + 図形.Height / 2
            Else
                MsgBox "指定された数字で入力して下さい。"
                Exit Sub
            End If

            xe = 配列(0)
            ye = 配列(1)
            
            ActiveSheet.Shapes.AddLine x, y, xe, ye
            
        Next 図形
    
    Else
        MsgBox "クリップボードにデータがありません！"
        
    End If
    
End Sub