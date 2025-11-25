Sub 選択したシートをCSVで保存する()
    Dim フォルダ名 As String
    Dim mySheet As Object
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "フォルダを選択して下さい"
        
        If .Show = -1 Then
            フォルダ名 = .SelectedItems(1)
        Else
            Exit Sub
        End If
    
    End With
    
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    
    For Each mySheet In ActiveWindow.SelectedSheets '選択したシートに対してループ処理を行う
        mySheet.Copy
        ActiveWorkbook.SaveAs Filename:=フォルダ名 & "\" & mySheet.Name & ".csv", FileFormat:=xlCSV, Local:=True
        ActiveWorkbook.Close False
        
    Next mySheet

End Sub