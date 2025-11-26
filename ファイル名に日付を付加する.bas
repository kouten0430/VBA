Sub ファイル名に日付を付加する()
    Dim 複数ファイル名 As Variant
    Dim ファイル名 As Variant
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    複数ファイル名 = Application.GetOpenFilename(Title:="ファイルを選択（複数選択可）", MultiSelect:=True)
    If TypeName(複数ファイル名) = "Boolean" Then  'キャンセルを押された場合の処理
        Exit Sub
    End If

    For Each ファイル名 In 複数ファイル名
        FSO.GetFile(ファイル名).Name = FSO.GetBaseName(ファイル名) & "_" & Format(Date, "yyyymmdd") & "." & FSO.GetExtensionName(ファイル名)

    Next ファイル名

End Sub