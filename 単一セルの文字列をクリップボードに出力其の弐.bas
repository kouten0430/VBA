Sub 単一セルの文字列をクリップボードに出力其の弐()
    'LfをCrLfに置換する
    Dim V As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

    V = Replace(ActiveCell.Value, vbLf, vbCrLf)

    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する
    
End Sub