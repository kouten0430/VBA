Sub 設備停止CSVから当該月が開始と終了にならない行を削除する()
    Dim 指定年月 As String
    Dim 開始日時の列番号 As Long
    Dim 終了日時の列番号 As Long
    Dim 行始 As Long
    Dim 行終 As Long
    Dim i As Long
    Dim 開始年月 As String
    Dim 終了年月 As String
    Dim myRange As Range
    Dim cnt As Long
    
    指定年月 = InputBox("削除しない月を西暦＋月で入力（例：202010）", , Year(Date) & Month(Date) + 1)
    If 指定年月 = "" Then Exit Sub
    
    開始日時の列番号 = Cells.Find("要求期間　開始日時", LookAt:=xlWhole).Column
    終了日時の列番号 = Cells.Find("要求期間　終了日時", LookAt:=xlWhole).Column
    行始 = Cells.Find("要求期間　開始日時", LookAt:=xlWhole).Row
    行終 = Cells(Rows.Count, 開始日時の列番号).End(xlUp).Row
    
    Application.ScreenUpdating = False  '画面表示の更新をオフにする

    For i = 行始 + 1 To 行終
        If Rows(i).Hidden = True Then GoTo continue '非表示行は処理対象としない
        
        開始年月 = Year(Cells(i, 開始日時の列番号).Value) & Month(Cells(i, 開始日時の列番号).Value)
        終了年月 = Year(Cells(i, 終了日時の列番号).Value) & Month(Cells(i, 終了日時の列番号).Value)
        
        If 開始年月 <> 指定年月 And 終了年月 <> 指定年月 Then
            If myRange Is Nothing Then
                Set myRange = Cells(i, 開始日時の列番号)    '条件に一致した一番最初のRange
            Else
                Set myRange = Union(myRange, Cells(i, 開始日時の列番号))    '条件に一致したRange
            End If
            
            cnt = cnt + 1
            
        End If
        
continue:

    Next i
        
    If Not myRange Is Nothing Then
        myRange.EntireRow.Delete '条件に一致した行をすべて削除する
    End If
    
    Application.ScreenUpdating = True  '画面表示の更新をオンにする
    
    MsgBox cnt & " 件削除しました。"

End Sub