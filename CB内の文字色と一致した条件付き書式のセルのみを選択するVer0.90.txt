Sub CB内の文字色と一致した条件付き書式のセルのみを選択する()
    Dim myRange As Range
    Dim myUni As Range
    Dim 色 As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    
    myLib.GetFromClipboard
    
    On Error Resume Next
    
    色 = myLib.GetText
    
    On Error GoTo 0

    If 色 <> "" Then
        For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
            If myRange.DisplayFormat.Font.Color = 色 And myRange.Address = myRange.MergeArea(1).Address Then
                If myUni Is Nothing Then
                    Set myUni = myRange    '検索に一致した一番最初のセル
                Else
                    Set myUni = Union(myUni, myRange)    '検索に一致した二番目以降のセル
                End If
            End If
        Next myRange
        
        If Not myUni Is Nothing Then
            myUni.Select    '検索に一致したセルをまとめて選択する
        End If
        
    Else
        MsgBox "クリップボードにデータがありません！"
        
    End If

End Sub