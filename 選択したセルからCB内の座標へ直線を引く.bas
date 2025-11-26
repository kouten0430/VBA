Sub 選択したセルからCB内の座標へ直線を引く()
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    Dim CB As String
    Dim 配列 As Variant
    Dim tmp As String
    Dim myRange As Range
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
        
        tmp = InputBox("セルのどの位置を始点とするか選んで下さい。" & vbCrLf & vbCrLf & _
        "1:左上" & vbCrLf & _
        "2:上中心" & vbCrLf & _
        "3:右上" & vbCrLf & _
        "4:左中心" & vbCrLf & _
        "5:右中心" & vbCrLf & _
        "6:左下" & vbCrLf & _
        "7:下中心" & vbCrLf & _
        "8:右下" & vbCrLf & _
        "9:中心")
        
        For Each myRange In Selection
            If tmp = "" Then
                Exit Sub
            ElseIf tmp = "1" Then
                x = myRange.Left
                y = myRange.Top
            ElseIf tmp = "2" Then
                x = myRange.Left + myRange.Width / 2
                y = myRange.Top
            ElseIf tmp = "3" Then
                x = myRange.Left + myRange.Width
                y = myRange.Top
            ElseIf tmp = "4" Then
                x = myRange.Left
                y = myRange.Top + myRange.Height / 2
            ElseIf tmp = "5" Then
                x = myRange.Left + myRange.Width
                y = myRange.Top + myRange.Height / 2
            ElseIf tmp = "6" Then
                x = myRange.Left
                y = myRange.Top + myRange.Height
            ElseIf tmp = "7" Then
                x = myRange.Left + myRange.Width / 2
                y = myRange.Top + myRange.Height
            ElseIf tmp = "8" Then
                x = myRange.Left + myRange.Width
                y = myRange.Top + myRange.Height
            ElseIf tmp = "9" Then
                x = myRange.Left + myRange.Width / 2
                y = myRange.Top + myRange.Height / 2
            Else
                MsgBox "指定された数字で入力して下さい。"
                Exit Sub
            End If

            xe = 配列(0)
            ye = 配列(1)
            
            ActiveSheet.Shapes.AddLine x, y, xe, ye
            
        Next myRange
    
    Else
        MsgBox "クリップボードにデータがありません！"
        
    End If
    
End Sub