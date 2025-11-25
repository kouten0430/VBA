Sub 垂直方向に離れたセルのデータをCBに格納する行検索版()
    '選択したセルから垂直方向に離れたセルのデータを改行区切りでCBに格納します
    '選択したセルを黄色に塗りつぶします
    'ローカル的なマクロです
    Dim myRange As Range
    Dim 行 As Long
    Dim 列 As Long
    Dim tmp As String
    Dim V As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    For Each myRange In Selection
    
        On Error GoTo エラー処理
        
        行 = Range("C1:C1048576").Find(myRange.Value, LookAt:=xlWhole).Row
        
        On Error GoTo 0
        
        列 = myRange.Column
        tmp = Replace(myRange.Value, "　", "") & "　" & Replace(Cells(行, 列).Value, "--------" & vbLf, "") 'ここはお好みに合わせてカスタマイズ
        V = V & tmp & vbCrLf
        myRange.Interior.Color = 65535
    Next myRange
        
    V = Left(V, Len(V) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）
    
    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する
    
    Exit Sub
    
エラー処理:
    MsgBox "一致する所名はありません。"
End Sub