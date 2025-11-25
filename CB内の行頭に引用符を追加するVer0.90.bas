Sub CB内の行頭に引用符を追加する()
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    Dim 入力 As String
    Dim 配列 As Variant
    Dim i As Integer
    Dim 出力 As String

    myLib.GetFromClipboard

    On Error Resume Next

    入力 = myLib.GetText

    On Error GoTo 0

    If 入力 <> "" Then
        配列 = Split(入力, vbCrLf)

        i = 0

        Do While i <= UBound(配列)
            出力 = 出力 & ">" & 配列(i) & vbCrLf

            i = i + 1

        Loop

        出力 = Left(出力, Len(出力) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）

        myLib.SetText 出力  '変数の値をDataObjectに格納する
        myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する

    Else
        MsgBox "クリップボードにデータがありません！"

    End If

End Sub