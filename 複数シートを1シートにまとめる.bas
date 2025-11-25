Sub 複数シートを1シートにまとめる()
    '「Ctrl」キー（または「Shift」キー）を押しながら、複数選択したシートを１シートにまとめます。
    'テンプレートを元にして新規シートを１枚自動作成し、選択されているシートのデータを下方向に挿入していきます。
    'まとめるシートはすべて同じ様式であることを想定しています。
    Dim mySWs As Sheets
    Dim myTWs As Worksheet
    Dim myNWs As Worksheet
    Dim yr As Range
    Dim y As Long
    Dim x As Integer
    Dim myWs As Worksheet
    Dim cye As Long
    Dim pye As Long
    Dim ctp As Integer
    Dim ct As Integer
    
retry:
        On Error Resume Next
    
        Set yr = Nothing 'retryした時のためのリセット

        Set yr = Application.InputBox(Prompt:="テンプレートとなるシートで、見出しの1行下を選択し、" & vbCrLf & _
        "OKして下さい。", Type:=8)
            If yr Is Nothing Then    'キャンセルを押された場合の処理
                Exit Sub
            ElseIf yr.Resize(1, 1).Value = "" Then  '複数範囲を選択された場合も想定して、Resizeを適用する
                MsgBox "空白以外のセルを選択してください。"
                GoTo retry
            End If

        On Error GoTo 0
        
    
    Set mySWs = ActiveWindow.SelectedSheets 'マクロ実行前に選択されているシート
    
    Set myTWs = Worksheets(yr.Parent.Name) 'テンプレートとなるシート
    
    myTWs.Copy Before:=Worksheets(1)
    Set myNWs = Worksheets(1)   'テンプレートを元に新規作成されたシート
    
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    
    y = yr.Row  'yrから行を取得
    x = yr.Column   'yrから列を取得

    For Each myWs In mySWs
        If myTWs.Name <> myWs.Name And myWs.ProtectContents = False Then  'テンプレートとなるシート、保護されたシート以外を処理する
            If Not myWs.AutoFilter Is Nothing Then   'コピーするシートにオートフィルターが設定されている場合は解除する
                myWs.Cells(y, x).AutoFilter
            End If
        
            cye = myWs.Cells(y, x).CurrentRegion.Rows(myWs.Cells(y, x).CurrentRegion.Rows.Count).Row
            pye = myNWs.Cells(y, x).CurrentRegion.Rows(myNWs.Cells(y, x).CurrentRegion.Rows.Count).Row
        
            myWs.Rows(y & ":" & cye).Copy
            myNWs.Rows(pye + 1).Insert
            myWs.Application.CutCopyMode = False  'コピーモードを解除する
        
        ElseIf myWs.ProtectContents Then
            ctp = ctp + 1

        End If

        ct = ct + 1
        Application.StatusBar = "処理実行中．．．" & ct & "/" & mySWs.Count
        
    Next myWs
    
    
    MsgBox "処理成功：" & ct - ctp & "/" & mySWs.Count & vbCrLf & "保護シート：" & ctp
    
    Application.StatusBar = False

End Sub