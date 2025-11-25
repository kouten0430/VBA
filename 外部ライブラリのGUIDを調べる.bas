Sub 外部ライブラリのGUIDを調べる()
'Microsoft Visual Basic for Applications Extensibilityを参照設定して下さい
'「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェックを入れて下さい
'ｱｸﾃｨﾌﾞｼｰﾄのA～D列に参照中のﾗｲﾌﾞﾗﾘの名称、GUID、ﾒｼﾞｬｰﾊﾞｰｼﾞｮﾝ、ﾏｲﾅｰﾊﾞｰｼﾞｮﾝを出力します
'個人用マクロブック以外のアクティブブックを対象にします
Dim myRef As Variant
Dim i As Integer
    i = 1
    Cells(i, 1).Value = "Name"
    Cells(i, 2).Value = "GUID"
    Cells(i, 3).Value = "Major"
    Cells(i, 4).Value = "Minor"
For Each myRef In ActiveWorkbook.VBProject.References
    i = i + 1
    Cells(i, 1).Value = myRef.Name
    Cells(i, 2).Value = myRef.GUID
    Cells(i, 3).Value = myRef.Major
    Cells(i, 4).Value = myRef.Minor
Next
End Sub