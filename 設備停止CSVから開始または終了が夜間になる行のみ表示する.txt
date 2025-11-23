Sub 設備停止CSVから開始または終了が夜間になる行のみ表示する()
    Dim 開始日時の列番号 As Long
    Dim 終了日時の列番号 As Long
    Dim 行始 As Long
    Dim 行終 As Long
    Dim i As Long
    Dim 開始日 As Date
    Dim 終了日 As Date
    Dim myRange As Range
    Dim cnt As Long
    
    開始日時の列番号 = Cells.Find("要求期間　開始日時", LookAt:=xlWhole).Column
    終了日時の列番号 = Cells.Find("要求期間　終了日時", LookAt:=xlWhole).Column
    行始 = Cells.Find("要求期間　開始日時", LookAt:=xlWhole).Row
    行終 = Cells(Rows.Count, 開始日時の列番号).End(xlUp).Row
    
    Application.ScreenUpdating = False  '画面表示の更新をオフにする

    For i = 行始 + 1 To 行終
        If Rows(i).Hidden = True Then GoTo continue '非表示行は処理対象としない
        
        開始日 = Cells(i, 開始日時の列番号).Value - Fix(Cells(i, 開始日時の列番号).Value)  '整数部分を取り除き時分のみにする
        終了日 = Cells(i, 終了日時の列番号).Value - Fix(Cells(i, 終了日時の列番号).Value) '整数部分を取り除き時分のみにする
        
        If 0.875 <= 開始日 Or 開始日 <= 0.291655092592593 Or 0.875 <= 終了日 Or 終了日 <= 0.291655092592593 Then '21時～6時59分59秒の場合
        
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