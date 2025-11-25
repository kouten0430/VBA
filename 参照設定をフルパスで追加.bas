Sub 参照設定をフルパスで追加()
'「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェックを入れて下さい
'個人用マクロブック以外のアクティブブックを対象にします
    Dim myRef As Variant
    Set myRef = ActiveWorkbook.VBProject.References.AddFromFile("C:\Program Files (x86)\Microsoft Office\root\VFS\SystemX86\FM20.DLL")
End Sub