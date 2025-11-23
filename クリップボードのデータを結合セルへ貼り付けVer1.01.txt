Sub クリップボードのデータを結合セルへ貼り付け()
    'Microsoft Forms 2.0 Object Libraryを参照設定して下さい

    Dim Dobj As DataObject
    Dim V As Variant
    Dim i As Integer
    Dim Y As Integer
    Dim X As Integer
    
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
        Y = ActiveCell.Row
        X = ActiveCell.Column

        Do While i <= UBound(V)
            If Cells(Y, X).Address = Cells(Y, X).MergeArea(1).Address _
            And Rows(Y).Hidden = False Then
                A = CStr(V(i))
                Cells(Y, X).Value = A
                Y = Y + 1
                i = i + 1
            Else
                Y = Y + 1
            End If
        Loop
    Else
        MsgBox "クリップボードにデータがありません！"
    End If

End Sub