Sub 指定文字以外を半角にする()
    '文字列を半角にしたいセルを選択（複数可）した状態で実行して下さい
    '指定文字は全角で指定して下さい
    '指定文字はセル参照とすることもできます
    Dim myRange As Range
    Dim i As Integer
    Dim n As Integer
    Dim V(99) As String
    
    For i = 0 To 99
        V(i) = Application.InputBox(Prompt:="半角にしない文字列を指定して下さい。" & vbCrLf & "（InputBoxが連続で表示され、複数指定できます。" _
        & vbCrLf & "これ以上指定がない場合は、ブランクでOKして下さい）", Type:=2)
    
        If V(i) = "" Then
            i = i - 1   '添え字の最大値からブランク分を除く
            Exit For
        ElseIf V(i) = "False" Then
            Exit Sub
        End If
    
    Next i
    
    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
        myRange.Value = StrConv(myRange.Value, vbNarrow)
        
        For n = 0 To i  '指定した文字列をすべて全角に復元する
            myRange.Value = Replace(myRange.Value, StrConv(V(n), vbNarrow), V(n))
        Next n

    Next myRange
    
End Sub