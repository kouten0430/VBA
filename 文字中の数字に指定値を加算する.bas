Sub 文字中の数字に指定値を加算する()
    '選択範囲に対して処理を行います
    '数字は半角・全角どちらでも処理可（漢数字は処理不可）
    '処理対象の数字の桁数を維持します
    Dim 零 As String
    Dim ns As Integer
    Dim ne As Integer
    Dim 値 As Long
    Dim tmp As Integer
    Dim myRange As Range
    Dim 数値 As Long
    Dim 数字 As String
    Dim 増量 As Long
    
    零 = "000000000"
    
    ns = InputBox("左から何文字目を開始位置にしますか？")
    ne = InputBox("開始位置から何文字の数字を対象にしますか？")
    値 = InputBox("加算する値を入力して下さい。" & vbCrLf & "（マイナスの値を入力すると減算になります）")
    tmp = MsgBox("入力した値で選択順にインクリメントしますか？", vbYesNo + vbDefaultButton2)

    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)
        If myRange.Value <> "" And TypeName(myRange.Value) <> "Date" Then   'セルの値が空白,日付の場合は処理をしない
            数値 = Mid(myRange.Value, ns, ne)
            数値 = 数値 + 値 + 増量
            
            If 数値 < 0 Then 数値 = 0  '減算した値が0未満の場合は0を下限とする
            
            If tmp = vbYes Then 増量 = 増量 + 値    'インクリメントする場合の処理
            
            If Mid(myRange.Value, ns, ne) = StrConv(Mid(myRange.Value, ns, ne), vbWide) Then
                数字 = StrConv(Right(零 & 数値, ne), vbWide)
            Else
                数字 = Right(零 & 数値, ne)
            End If
            
            myRange.Value = Application.WorksheetFunction.Replace(myRange.Value, ns, ne, 数字)
            
        End If
        
    Next myRange
    
End Sub