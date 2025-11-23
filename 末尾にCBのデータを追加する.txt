Sub 末尾にCBのデータを追加する()
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    Dim 全文字列 As String
    Dim 分割文字列 As Variant
    Dim i As Integer
    Dim myRange As Range
    
    myLib.GetFromClipboard
    
    On Error Resume Next
    
    全文字列 = myLib.GetText
    
    On Error GoTo 0
    
    If 全文字列 <> "" Then
        分割文字列 = Split(全文字列, vbCrLf)  '全文字列を改行で分割し配列に格納する
        i = 0
    
        For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
            If i <= UBound(分割文字列) Then
                If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルのみ処理する
                    myRange.Value = myRange.Value & 分割文字列(i)
                    i = i + 1
                End If
            Else
                Exit For
            End If

        Next myRange
        
    Else
        MsgBox "クリップボードにデータがありません！"

    End If

End Sub