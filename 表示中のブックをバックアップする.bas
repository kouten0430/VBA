Sub 表示中のブックをバックアップする()
    '最後に保存された状態をバックアップします
    Dim コピー先 As String
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    コピー先 = "C:\Users\wrfmf\Documents\雑庫\新しいフォルダー\"    '必要に応じて変更する
    
    FSO.GetFile(ActiveWorkbook.FullName).Copy コピー先 & Format(Now, "yyyymmddhhmmss") & "." & FSO.GetExtensionName(ActiveWorkbook.FullName), False
    
End Sub