Sub コマンドバーを追加()
    CommandBars.Add Name:="★★★"
    CommandBars("★★★").Visible = True    '追加したコマンドバーは規定で非表示になっているため、手動で表示させる必要がある
End Sub
Sub コマンドバーを削除()
    CommandBars("★★★").Delete
End Sub
Sub コマンドバーを一時的に追加()
    CommandBars.Add Name:="★★★", Temporary:=True
End Sub
Sub ★★★コマンドバーコントロールを追加()
    Dim myButton As CommandBarButton

    Set myButton = CommandBars("★★★").Controls.Add
    myButton.Caption = "☆☆☆" '追加するメニュー名を入力
    myButton.OnAction = "☆☆☆"    '実行するマクロ名を入力
    myButton.Style = msoButtonCaption   'メニューの表示形式を設定する
End Sub
Sub ★★★コマンドバーコントロールを削除()
    CommandBars("★★★").Controls("☆☆☆").Delete   '削除するメニュー名を入力
End Sub
Sub Cellコマンドバーコントロールを追加()
    Dim myButton As CommandBarButton

    Set myButton = CommandBars("Cell").Controls.Add(Before:=1)  'Before:=で配置位置を指定する。省略で末尾に配置
    myButton.Caption = "★★★" '追加するメニュー名を入力
    myButton.OnAction = "★★★"    '実行するマクロ名を入力
    myButton.BeginGroup = False  'Trueで上側に区切り線が表示される。Falseで区切り線なし
    
End Sub
Sub Cellサブメニューを含むコマンドバーコントロールを追加()
    Dim myButton As CommandBarPopup

    Set myButton = CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Before:=1) 'Before:=で配置位置を指定する。省略で末尾に配置
    myButton.Caption = "★★★" '追加するメニュー名を入力
    myButton.BeginGroup = False  'Trueで上側に区切り線が表示される。Falseで区切り線なし
    
End Sub
Sub Cellサブメニューを追加()
    Dim myButton As CommandBarButton

    Set myButton = CommandBars("Cell").Controls("★★★").Controls.Add(Before:=1)  'Before:=で配置位置を指定する。省略で末尾に配置
    myButton.Caption = "☆☆☆" '追加するメニュー名を入力
    myButton.OnAction = "☆☆☆"    '実行するマクロ名を入力
    myButton.BeginGroup = False  'Trueで上側に区切り線が表示される。Falseで区切り線なし
    
End Sub
Sub Cellコマンドバーコントロールを削除()
    CommandBars("Cell").Controls("★★★").Delete   '削除するメニュー名を入力
End Sub
Sub Cellサブメニューを削除()
    CommandBars("Cell").Controls("★★★").Controls("☆☆☆").Delete   '削除するサブメニュー名を入力
End Sub
Sub Cellコマンドバーコントロールをリセット()
    CommandBars("Cell").Reset
End Sub