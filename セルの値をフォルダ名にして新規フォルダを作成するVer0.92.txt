Sub リストを使用して新規フォルダを作成する()
    Dim mb As Integer
    Dim Dn As String
    Dim myRange As Range
    Dim ct As Long
    
    mb = MsgBox(prompt:="選択中のセルの文字列をフォルダ名にして新規フォルダを作成します。" & vbCrLf & _
    "良ければOKし、次に新規フォルダを作成する場所を選択して下さい。", Buttons:=vbOKCancel)
        If mb = 2 Then  '「キャンセル」を選択した場合はプロシージャを終了する
            Exit Sub
        End If
            
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "フォルダを選択して下さい"
        
        If .Show = -1 Then
            Dn = .SelectedItems(1)
        Else
            Exit Sub
        End If
    
    End With
    
    For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上の値のみ取り出す
            If myRange.MergeArea(1).Value <> "" Then 'セルの値が空白以外の場合のみ処理を行う
                If Dir(Dn & "\" & CStr(myRange.MergeArea(1).Value), vbDirectory) = "" Then '同名フォルダが存在しない場合のみ処理を行う
                    MkDir Path:=Dn & "\" & CStr(myRange.MergeArea(1).Value) '新規フォルダを作成する
                    ct = ct + 1
                End If
            End If
        End If
    Next myRange
    
    MsgBox "作成成功：" & ct & " フォルダ"

End Sub