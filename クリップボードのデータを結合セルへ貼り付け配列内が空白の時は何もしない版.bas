Sub クリップボードのデータを結合セルへ貼り付け配列内が空白の時は何もしない版()
    'Microsoft Forms 2.0 Object Libraryを参照設定して下さい

    Dim Dobj As DataObject
    Dim V As Variant
    Dim i As Integer
    Dim y As Integer
    Dim x As Integer
    
    Set Dobj = New DataObject
    With Dobj
        .GetFromClipboard
        On Error Resume Next
        V = .GetText
        On Error GoTo 0
    End With
    
    If Not IsEmpty(V) Then
        V = Split(CStr(V), vbCrLf)
        i = 0
        y = ActiveCell.Row
        x = ActiveCell.Column

        Do While i <= UBound(V)
            If Cells(y, x).Address = Cells(y, x).MergeArea(1).Address _
            And Rows(y).Hidden = False Then
                a = CStr(V(i))
                
                If a <> "" Then '配列内が空白の時は何もしない
                    Cells(y, x).Value = a
                End If
                
                y = y + 1
                i = i + 1
            Else
                y = y + 1
            End If
        Loop
    Else
        MsgBox "クリップボードにデータがありません！"
    End If

End Sub