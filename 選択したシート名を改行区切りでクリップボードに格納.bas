Sub 選択したシート名を改行区切りでクリップボードに格納()
    '選択したシートは左側から順次ループ処理されます
    Dim V As String
    Dim mySheet As Object
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する

    For Each mySheet In ActiveWindow.SelectedSheets '選択したシートに対してループ処理を行う
        V = V & mySheet.Name & vbCrLf
    Next mySheet
    
    V = Left(V, Len(V) - 2) '最後の改行区切りを取り除く（CrLfは2文字）

    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する

End Sub