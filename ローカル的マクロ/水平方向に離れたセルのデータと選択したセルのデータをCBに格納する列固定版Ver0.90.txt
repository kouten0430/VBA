Sub 水平方向に離れたセルのデータと選択したセルのデータをCBに格納する列固定版()
    '水平方向に離れたセルのデータと選択したセルのデータを合わせ改行区切りでCBに格納します
    '選択したセルを青色に塗りつぶします
    'ローカル的なマクロです
    Dim 列 As Long
    Dim myRange As Range
    Dim V As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    列 = 2
    
    For Each myRange In Selection
        V = V & Cells(myRange.Row, 列).Value & "," & myRange.Value & vbCrLf '水平方向に離れたセルのデータと選択したセルのデータを区切る文字列はお好みで
        myRange.Interior.Color = 15773696
    Next myRange
        
    V = Left(V, Len(V) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）
    
    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する

End Sub