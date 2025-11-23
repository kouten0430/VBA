Sub CB内データのいずれかと一致するセルをClearContentsする()
    'CB内のデータは改行区切りになっている必要があります
    'あらかじめ処理範囲を選択してから実行して下さい
    'ローカル的なマクロです
    Dim V As Variant
    Dim myRange As Range
    Dim i As Integer
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    
    myLib.GetFromClipboard
    
    On Error Resume Next
    
    V = myLib.GetText
    
    On Error GoTo 0
        
    If Not IsEmpty(V) Then
        V = Split(CStr(V), vbCrLf)

        For Each myRange In Selection

            i = 0

            Do While i <= UBound(V)
                If CStr(myRange.Value) = V(i) And V(i) <> "" Then
                    myRange.ClearContents
                    Exit Do   '残りの検索をスキップ
                    
                Else
                    i = i + 1
                End If
            Loop

        Next myRange

    Else
        MsgBox "クリップボードにデータがありません！"
        
    End If

End Sub