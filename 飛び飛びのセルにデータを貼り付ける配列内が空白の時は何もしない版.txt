Sub 飛び飛びのセルにデータを貼り付ける配列内が空白の時は何もしない版()
    'クリップボード内のデータを選択した飛び飛びのセルに貼り付けすることができます
    'データは改行区切りとなっている必要があります
    '結合セルを単一セルのように扱うことができます
    Dim V As Variant
    Dim i As Integer
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
        
        If Selection.Count > 1 Then
            For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
                If i <= UBound(V) Then
                    If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルにのみ処理を行う
                        If CStr(V(i)) <> "" Then    '配列内が空白の時は何もしない
                            myRange.Value = CStr(V(i))
                        End If
                        
                        i = i + 1   '配列の次の添え字を作成
                        
                    End If
                Else
                    Exit For    '配列の添え字が最大値を超えたらFor Eachを抜ける
                End If
        
            Next myRange
            
        Else
            ActiveCell.Value = CStr(V(0))
            
        End If

    Else
        MsgBox "クリップボードにデータがありません！"

    End If
    
End Sub