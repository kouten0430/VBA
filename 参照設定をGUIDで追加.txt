Sub 参照設定をGUIDで追加()
'「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェックを入れて下さい
'個人用マクロブック以外のアクティブブックを対象にします
    Dim myRef As Variant
    Set myRef = ActiveWorkbook.VBProject.References.AddFromGuid("{0D452EE1-E08F-101A-852E-02608C4D0BB4}", 2, 0)
End Sub