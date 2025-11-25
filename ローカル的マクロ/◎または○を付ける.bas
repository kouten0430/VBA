Sub ◎または○を付ける()
    '丸を付ける行（どこでもいい）を選択して実行します
    '名前の行と配列の内容を比較し、一致する列に◎または○を付けます
    'ローカル的なマクロです
    Dim 名前の行 As Long
    Dim 現在の行 As Long
    Dim V As Variant
    Dim i As Integer
    Dim 列 As Integer
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    名前の行 = 3
    現在の行 = Selection.Row
    
    myLib.GetFromClipboard
    
    On Error Resume Next
    
    V = myLib.GetText
    
    On Error GoTo 0

    If Not IsEmpty(V) Then
        V = Split(CStr(V), vbCrLf)
        
        For i = 0 To UBound(V)
            On Error GoTo エラー処理
            
            列 = Range(Cells(名前の行, 1), Cells(名前の行, 99)).Find(V(i), LookAt:=xlWhole).Column
            
            On Error GoTo 0
            
            If i = 0 Then
                Cells(現在の行, 列).Value = "◎"
            Else
                Cells(現在の行, 列).Value = "○"
            End If
            
        Next i

    Else
        MsgBox "クリップボードにデータがありません！"
        
    End If
    
    Exit Sub

エラー処理:
    MsgBox V(i) & " が、どれとも一致しないので処理を中断しました。"
    
End Sub