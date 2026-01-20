Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim filePath As String
    filePath = ThisWorkbook.FullName

    ' 属性が読み取り専用でない場合のみ実行
    If (GetAttr(filePath) And vbReadOnly) = 0 Then
        
        ' --- 追加部分：保存の確認 ---
        If Not ThisWorkbook.Saved Then
            If MsgBox("'" & ThisWorkbook.Name & "' への変更を保存しますか?", _
                      vbYesNo + vbExclamation, "Microsoft Excel") = vbYes Then
                ThisWorkbook.Save
            End If
        End If
        ' ----------------------------

        ' 1. ファイル属性を読み取り専用に設定
        SetAttr filePath, vbReadOnly
        
        ' 2. 「保存済み」の状態だとExcelに誤認させて、確認ダイアログを抑制する
        ThisWorkbook.Saved = True
    End If
End Sub