Sub 選択範囲をオリジナルの順番で並び替える()
    'Microsoft Forms 2.0 Object Libraryを参照設定して下さい
    'あらかじめ、オリジナルの順番をカンマ区切りでクリップボードに格納しておいて下さい
    '範囲選択後、Tabキーでアクティブセル（白抜き）を並び替えのキーとなる列に移動して下さい
    Dim V As String
    
    Set Dobj = New DataObject
    With Dobj
        .GetFromClipboard
        On Error Resume Next
        V = .GetText
        On Error GoTo 0
    End With
    
    If V <> Empty Then
        With ActiveSheet.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ActiveCell, CustomOrder:="""," & V & """"
            .SetRange Selection
            .Header = xlNo
            .Apply
        End With
    Else
        MsgBox "クリップボードにデータがありません！"
    End If
End Sub