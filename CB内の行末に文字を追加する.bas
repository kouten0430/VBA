Sub CB内の行末に文字を追加する()
    '行末のスペース（全角・半角）は無視されます
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    Dim 入力 As String
    Dim 指定文字 As String
    Dim 配列 As Variant
    Dim i As Integer
    Dim 逆 As String
    Dim k As Integer
    Dim 出力 As String

    myLib.GetFromClipboard

    On Error Resume Next

    入力 = myLib.GetText

    On Error GoTo 0

    If 入力 <> "" Then
        指定文字 = StrReverse(InputBox("指定文字を入力"))
        If 指定文字 = "" Then Exit Sub

        配列 = Split(入力, vbCrLf)

        i = 0

        Do While i <= UBound(配列)
            逆 = StrReverse(配列(i))

            For k = 1 To Len(逆)
                If Mid(逆, k, 1) Like "[! 　]" Then
                    出力 = 出力 & StrReverse(Application.WorksheetFunction.Replace(逆, k, 0, 指定文字)) & vbCrLf
                    GoTo skip

                End If

            Next k

            出力 = 出力 & 配列(i) & vbCrLf

skip:

            i = i + 1

        Loop

        出力 = Left(出力, Len(出力) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）

        myLib.SetText 出力  '変数の値をDataObjectに格納する
        myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する

    Else
        MsgBox "クリップボードにデータがありません！"

    End If

End Sub