Sub 同じ文字を下方向へnセル分貼り付ける()
    'Microsoft Forms 2.0 Object Libraryを参照設定して下さい
    'クリップボードのデータを丸ごと1セルに貼り付ける処理をn回繰り返します
    '結合セルは1セルとしてカウントします
    '非表示セルは1セルとしてカウントしません（つまり可視セルのみに貼り付け）

    Dim Dobj As DataObject
    Dim V As Variant
    Dim i As Integer
    Dim Y As Integer
    Dim X As Integer
    Dim YE As Variant
    
    YE = Application.InputBox(Prompt:="下方向へ何セル分貼り付けますか？", Type:=1)
        If TypeName(YE) = "Boolean" Then
            Exit Sub
        End If
    
    Set Dobj = New DataObject
    With Dobj
        .GetFromClipboard
        On Error Resume Next
        V = .GetText
        On Error GoTo 0
    End With
    
    If Not IsEmpty(V) Then
        i = 1
        Y = ActiveCell.Row
        X = ActiveCell.Column

        Do While i <= YE
            If Cells(Y, X).Address = Cells(Y, X).MergeArea(1).Address _
            And Rows(Y).Hidden = False Then
                Cells(Y, X).Value = V
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