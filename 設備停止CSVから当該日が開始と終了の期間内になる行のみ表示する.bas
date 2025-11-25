Sub 設備停止CSVから当該日が開始と終了の期間内になる行のみ表示する()
    Dim 指定日 As Date
    Dim 開始日時の列番号 As Long
    Dim 終了日時の列番号 As Long
    Dim 行始 As Long
    Dim 行終 As Long
    Dim i As Long
    Dim 開始日 As Date
    Dim 終了日 As Date
    Dim myRange As Range
    Dim cnt As Long
    
    指定日 = DateValue(InputBox("表示する日を西暦/月/日で入力（例：2021/6/26）", , Year(Date) & "/" & Month(Date) & "/"))
    
    開始日時の列番号 = Cells.Find("要求期間　開始日時", LookAt:=xlWhole).Column
    終了日時の列番号 = Cells.Find("要求期間　終了日時", LookAt:=xlWhole).Column
    行始 = Cells.Find("要求期間　開始日時", LookAt:=xlWhole).Row
    行終 = Cells(Rows.Count, 開始日時の列番号).End(xlUp).Row
    
    Application.ScreenUpdating = False  '画面表示の更新をオフにする

    For i = 行始 + 1 To 行終
        If Rows(i).Hidden = True Then GoTo continue '非表示行は処理対象としない
        
        開始日 = WorksheetFunction.Floor(Cells(i, 開始日時の列番号).Value, 1)   '時分は切り捨てる
        終了日 = WorksheetFunction.Floor(Cells(i, 終了日時の列番号).Value, 1)   '時分は切り捨てる
        
        If 開始日 <= 指定日 And 指定日 <= 終了日 Then
        
        Else
            If myRange Is Nothing Then
                Set myRange = Cells(i, 開始日時の列番号)    '条件に一致しなかった一番最初のRange
            Else
                Set myRange = Union(myRange, Cells(i, 開始日時の列番号))    '条件に一致しなかったRange
            End If
            
            cnt = cnt + 1
            
        End If
        
continue:

    Next i
        
    If Not myRange Is Nothing Then
        myRange.EntireRow.Hidden = True '条件に一致しなかった行をすべて非表示にする
    End If
    
    Application.ScreenUpdating = True  '画面表示の更新をオンにする
    
    MsgBox cnt & " 件非表示にしました。"

End Sub