Sub 給電事故情報ダウンロード()
    '給電事故・瞬低検索パラメータの入力（ダウンロード用）の画面を表示した状態で実行する
    'その他のIEは閉じておく
    '発行箇所・発生年度（開始年度）・給電事故または瞬停のチェックはあらかじめ入力しておく
    Dim ie As Object
    Dim sh As Object
    Dim win As Object
    Dim 開始年度 As Integer
    Dim 終了年度 As Integer
    
    Set sh = CreateObject("Shell.Application")
    
    For Each win In sh.Windows
        If win.Name = "Internet Explorer" Then
            Set ie = win
            Exit For
        End If
    Next

    開始年度 = ie.document.getElementsByName("txthnendo")(0).Value
    終了年度 = InputBox("終了年度を入力して下さい")
    
    SetForegroundWindow (ie.hWnd)   'IEを最前面に表示する
    
    Do While 開始年度 <= 終了年度
        ie.document.getElementsByName("txthnendo")(0).Value = 開始年度
        
        ie.document.getElementsByTagName("INPUT")(22).Click
        
        Sleep 10000
        
        SendKeys "%n", True 'Alt+N で通知バーにフォーカスがあたる
        SendKeys "{TAB}", True
        SendKeys "{ENTER}", True
        
        開始年度 = 開始年度 + 1
    
    Loop
    
    ie.Quit
        
End Sub