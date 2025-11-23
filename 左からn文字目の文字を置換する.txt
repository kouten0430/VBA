Sub 左からn文字目の文字を置換する()
    
    Dim myRange As Range
    Dim ns As Integer
    Dim ne As Integer
    Dim tm As String
    
    ns = InputBox("左から何文字目を置換開始位置にしますか？")
    ne = InputBox("開始位置から何文字置換しますか？" _
    & vbCrLf & "※文字を挿入する場合は0を入力下さい。" _
    & vbCrLf & "　開始位置の左側に挿入されます。")
    tm = InputBox("置換または挿入する文字を入力して下さい。" _
    & vbCrLf & "※空白のままOKにすると、開始位置から" _
    & vbCrLf & "　指定文字数が削除（空白に置換）されます。")

    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)
        If myRange.Value <> "" And TypeName(myRange.Value) <> "Date" Then   'セルの値が空白,日付の場合は処理をしない
            myRange.Value = Application.WorksheetFunction.Replace(myRange.Value, ns, ne, tm)
        End If
    Next myRange
    
End Sub