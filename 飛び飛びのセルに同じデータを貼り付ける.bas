Sub 飛び飛びのセルに同じデータを貼り付ける()
    'クリップボード内の同じデータを選択した飛び飛びのセルに貼り付けすることができます
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
        If Selection.Count > 1 Then
            For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
                If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルにのみ処理を行う
                    myRange = V
                End If
        
            Next myRange
            
        Else
            ActiveCell.Value = V
            
        End If

    Else
        MsgBox "クリップボードにデータがありません！"

    End If
    
End Sub