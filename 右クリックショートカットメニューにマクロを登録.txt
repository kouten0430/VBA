Sub 右クリックショートカットメニューにマクロを登録()
    'このマクロ自体をメニューまたはサブメニューから呼び出している場合は、該当のメニューまたはサブメニューの削除はできません
    
    Dim sn As Variant
    Dim ct As Variant
    Dim bf As Variant
    Dim oa As Variant
    Dim bg As Variant
    Dim cts As Variant
    Dim myButton As CommandBarButton
    Dim myPopup As CommandBarPopup
    
    sn = Application.InputBox(Prompt:="処理内容を選択して下さい" & vbCrLf & "1:メニューの追加" & _
    vbCrLf & "2:サブメニューを含むメニューの追加" & vbCrLf & "3:サブメニューの追加" & vbCrLf & _
    "4:メニューの削除" & vbCrLf & "5:サブメニューの削除" & vbCrLf & "6:初期状態に戻す（リセット）", Type:=1)
        Select Case sn
        
        Case 1
ctr:
            ct = Application.InputBox(Prompt:="メニュー名を入力して下さい", Type:=2)
                If TypeName(ct) = "Boolean" Then
                    Exit Sub
                ElseIf ct = "" Then
                    GoTo ctr
                End If
oar:
            oa = Application.InputBox(Prompt:="メニューに登録するマクロ名を入力して下さい", Type:=2)
                If TypeName(oa) = "Boolean" Then
                    Exit Sub
                ElseIf oa = "" Then
                    GoTo oar
                End If
bfr:
            bf = Application.InputBox(Prompt:="メニューを上から何番目に配置しますか？（0で末尾）", Type:=1)
                If TypeName(bf) = "Boolean" Then
                    Exit Sub
                ElseIf bf = 0 Then
                    Set myButton = CommandBars("Cell").Controls.Add
                ElseIf bf >= 1 And bf <= CommandBars("Cell").Controls.Count Then
                    Set myButton = CommandBars("Cell").Controls.Add(Before:=bf) 'Before:=で配置位置を指定する。省略で末尾に配置
                Else
                    MsgBox "0～" & CommandBars("Cell").Controls.Count & "の数値で入力して下さい"
                    GoTo bfr
                End If
            myButton.Caption = ct '追加するメニュー名を入力
            myButton.OnAction = oa    '実行するマクロ名を入力
bgr:
            bg = Application.InputBox(Prompt:="メニューの上側に区切り線を入れますか？" & vbCrLf & _
            "1:入れる" & vbCrLf & "2:入れない", Type:=1)
                If TypeName(bg) = "Boolean" Then
                    Exit Sub
                ElseIf bg = 1 Then
                    myButton.BeginGroup = True  'Trueで上側に区切り線が表示される。Falseで区切り線なし
                ElseIf bg = 2 Then
                    myButton.BeginGroup = False
                Else
                    MsgBox "1または2を入力して下さい！"
                    GoTo bgr
                End If
                
        Case 2
ctrr:
            ct = Application.InputBox(Prompt:="サブメニューを含むメニュー名を入力して下さい", Type:=2)
                If TypeName(ct) = "Boolean" Then
                    Exit Sub
                ElseIf ct = "" Then
                    GoTo ctrr
                End If
bfrr:
            bf = Application.InputBox(Prompt:="サブメニューを含むメニューを上から何番目に配置しますか？（0で末尾）", Type:=1)
                If TypeName(bf) = "Boolean" Then
                    Exit Sub
                ElseIf bf = 0 Then
                    Set myPopup = CommandBars("Cell").Controls.Add(Type:=msoControlPopup)
                ElseIf bf >= 1 And bf <= CommandBars("Cell").Controls.Count Then
                    Set myPopup = CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Before:=bf) 'Before:=で配置位置を指定する。省略で末尾に配置
                Else
                    MsgBox "0～" & CommandBars("Cell").Controls.Count & "の数値で入力して下さい"
                    GoTo bfrr
                End If
            myPopup.Caption = ct '追加するメニュー名を入力
bgrr:
            bg = Application.InputBox(Prompt:="サブメニューを含むメニューの上側に区切り線を入れますか？" & vbCrLf & _
            "1:入れる" & vbCrLf & "2:入れない", Type:=1)
                If TypeName(bg) = "Boolean" Then
                    Exit Sub
                ElseIf bg = 1 Then
                    myPopup.BeginGroup = True  'Trueで上側に区切り線が表示される。Falseで区切り線なし
                ElseIf bg = 2 Then
                    myPopup.BeginGroup = False
                Else
                    MsgBox "1または2を入力して下さい！"
                    GoTo bgrr
                End If

        Case 3
ctsw:
            On Error GoTo ErrorHandler
            cts = Application.InputBox(Prompt:="サブメニュー名を入力して下さい", Type:=2)
                If TypeName(cts) = "Boolean" Then
                    Exit Sub
                ElseIf cts = "" Then
                    GoTo ctsw
                End If
ctw:
            ct = Application.InputBox(Prompt:="サブメニューの上位階層のメニュー名を入力して下さい", Type:=2)
                If TypeName(ct) = "Boolean" Then
                    Exit Sub
                ElseIf ct = "" Then
                    GoTo ctw
                End If
oaw:
            oa = Application.InputBox(Prompt:="サブメニューに登録するマクロ名を入力して下さい", Type:=2)
                If TypeName(oa) = "Boolean" Then
                    Exit Sub
                ElseIf oa = "" Then
                    GoTo oaw
                End If
bfw:
            bf = Application.InputBox(Prompt:="サブメニューを上から何番目に配置しますか？（0で末尾）", Type:=1)
                If TypeName(bf) = "Boolean" Then
                    Exit Sub
                ElseIf bf = 0 Then
                    Set myButton = CommandBars("Cell").Controls(ct).Controls.Add
                ElseIf bf >= 1 And bf <= CommandBars("Cell").Controls(ct).Controls.Count Then
                    Set myButton = CommandBars("Cell").Controls(ct).Controls.Add(Before:=bf) 'Before:=で配置位置を指定する。省略で末尾に配置
                Else
                    MsgBox "0～" & CommandBars("Cell").Controls(ct).Controls.Count & "の数値で入力して下さい"
                    GoTo bfw
                End If
            myButton.Caption = cts '追加するサブメニュー名を入力
            myButton.OnAction = oa    '実行するマクロ名を入力
bgw:
            bg = Application.InputBox(Prompt:="サブメニューの上側に区切り線を入れますか？" & vbCrLf & _
            "1:入れる" & vbCrLf & "2:入れない", Type:=1)
                If TypeName(bg) = "Boolean" Then
                    Exit Sub
                ElseIf bg = 1 Then
                    myButton.BeginGroup = True  'Trueで上側に区切り線が表示される。Falseで区切り線なし
                ElseIf bg = 2 Then
                    myButton.BeginGroup = False
                Else
                    MsgBox "1または2を入力して下さい！"
                    GoTo bgw
                End If
                        
        Case 4
ctx:
            On Error GoTo ErrorHandler
            ct = Application.InputBox(Prompt:="削除するメニュー名を入力して下さい", Type:=2)
                If TypeName(ct) = "Boolean" Then
                    Exit Sub
                ElseIf ct = "" Then
                    GoTo ctx
                End If
            CommandBars("Cell").Controls(ct).Delete   '削除するメニュー名を入力
            
        Case 5
ctsx:
            On Error GoTo ErrorHandler
            cts = Application.InputBox(Prompt:="削除するサブメニュー名を入力して下さい", Type:=2)
                If TypeName(cts) = "Boolean" Then
                    Exit Sub
                ElseIf cts = "" Then
                    GoTo ctsx
                End If
ctsz:
            ct = Application.InputBox(Prompt:="サブメニューの上位階層のメニュー名を入力して下さい", Type:=2)
                If TypeName(ct) = "Boolean" Then
                    Exit Sub
                ElseIf ct = "" Then
                    GoTo ctsz
                End If
            CommandBars("Cell").Controls(ct).Controls(cts).Delete   '削除するサブメニュー名を入力
            
        Case 6
            CommandBars("Cell").Reset
        
        End Select
        
    Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 5
            MsgBox "入力された名称が間違っているか、存在しません"
    End Select
End Sub