Sub CB内の行頭に連番を振る()
    '行頭のスペース（全角・半角）は無視されます
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    Dim 入力 As String
    Dim cnt As Integer
    Dim tmp As Integer
    Dim 配列 As Variant
    Dim i As Integer
    Dim k As Integer
    Dim 数字 As String
    Dim 出力 As String

    myLib.GetFromClipboard

    On Error Resume Next

    入力 = myLib.GetText

    On Error GoTo 0

    If 入力 <> "" Then
        cnt = InputBox("開始番号を入力")

        tmp = MsgBox("全角にしますか？", vbYesNoCancel + vbDefaultButton2)
        If tmp = vbCancel Then Exit Sub

        配列 = Split(入力, vbCrLf)

        i = 0

        Do While i <= UBound(配列)
            For k = 1 To Len(配列(i))
                If Mid(配列(i), k, 1) Like "[! 　]" Then
                    If tmp = vbYes Then
                        数字 = StrConv(cnt, vbWide)
                    ElseIf tmp = vbNo Then
                        数字 = cnt
                    End If

                    出力 = 出力 & Application.WorksheetFunction.Replace(配列(i), k, 0, 数字) & vbCrLf

                    cnt = cnt + 1

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