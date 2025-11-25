Sub 飛び飛びのセルに改行して冒頭にデータを貼り付ける()
    'クリップボード内のデータを選択した飛び飛びのセルに貼り付けすることができます
    'データは改行区切りとなっている必要があります
    '結合セルを単一セルのように扱うことができます
    Dim V As Variant
    Dim i As Integer
    Dim a As String
    Dim myRange As Range
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する

    myLib.GetFromClipboard
        On Error Resume Next
    V = myLib.GetText
        On Error GoTo 0

    If Not IsEmpty(V) Then
        V = Split(CStr(V), vbCrLf)
        i = 0
        
        For Each myRange In Selection
            If i <= UBound(V) Then
                If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルにのみ処理を行う
                    a = CStr(V(i))
                    myRange.Value = a & vbLf & myRange.Value
                    i = i + 1   '配列の次の添え字を作成
                End If
            Else
                Exit For    '配列の添え字が最大値を超えたらFor Eachを抜ける
            End If
        
        Next myRange

    Else
        MsgBox "クリップボードにデータがありません！"

    End If
    
End Sub