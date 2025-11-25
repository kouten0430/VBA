Sub 左からn文字目の文字をCBのデータに置換する()
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    Dim 全文字列 As String
    Dim 分割文字列 As Variant
    Dim i As Integer
    Dim myRange As Range
    Dim ns As Integer
    Dim ne As Integer
    
    ns = InputBox("左から何文字目を置換開始位置にしますか？")
    ne = InputBox("開始位置から何文字置換しますか？" _
    & vbCrLf & "※文字を挿入する場合は0を入力下さい。" _
    & vbCrLf & "　開始位置の左側に挿入されます。")
    
    myLib.GetFromClipboard
    
    On Error Resume Next
    
    全文字列 = myLib.GetText
    
    On Error GoTo 0
    
    If 全文字列 <> "" Then
        分割文字列 = Split(全文字列, vbCrLf)  '全文字列を改行で分割し配列に格納する
        i = 0
    
        For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
            If i <= UBound(分割文字列) Then
                If myRange.Value <> "" And TypeName(myRange.Value) <> "Date" Then   'セルの値が空白,日付の場合は処理をしない
                    myRange.Value = Application.WorksheetFunction.Replace(myRange.Value, ns, ne, 分割文字列(i))
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