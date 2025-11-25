Sub 複数のセルの色で絞込みを行う()
    'クリップボードに格納されたセル色の値を参照し、OR条件で絞込みします
    '実行前に絞り込みを行う列範囲（見出しを除く）を選択しておきます
    'セル色が一致しない行を非表示にします（オートフィルターを使いません）
    Dim V As Variant
    Dim i As Integer
    Dim x As Integer
    Dim y As Long
    Dim Yn As Long
    Dim myRange As Range
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    
    myLib.GetFromClipboard
        On Error Resume Next
    V = myLib.GetText
        On Error GoTo 0
        
    If Not IsEmpty(V) Then
        V = Split(CStr(V), vbCrLf)
        x = Selection.Column

        For y = Selection.Row To Selection.Rows(Selection.Rows.Count).Row
            i = 0
        
            Do While i <= UBound(V)
                If CStr(Cells(y, x).Interior.Color) = V(i) Then '配列の内容と一致している場合は行を進める
                        Yn = y + 1
                        Do While Cells(y, x).Address = Cells(Yn, x).MergeArea(1).Address    '結合セルを抜けるまで行を進める
                            Yn = Yn + 1
                        Loop
                        y = Yn - 1
                    
                    GoTo nx
                Else
                    i = i + 1
                End If
            Loop
            
            If myRange Is Nothing Then
                Set myRange = Range(y & ":" & y)    '配列の内容全てと一致しなかった一番最初の行
            Else
                Set myRange = Union(myRange, Range(y & ":" & y))    '配列の内容全てと一致しなかった行
            End If
             
nx:
        Next y
        
        myRange.EntireRow.Hidden = True '検索に一致しなかった行をすべて非表示にする
    Else
        MsgBox "クリップボードにデータがありません！"
    End If

End Sub