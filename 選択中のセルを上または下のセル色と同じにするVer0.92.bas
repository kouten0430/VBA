Sub 選択中のセルを上または下のセル色と同じにする()
    '上または下の可視セルのセル色と同じにする
    '結合セルの左上のOffsetは結合エリアの外から始まる
    Dim tmp As Integer
    Dim myRange As Range
    Dim i As Long
    
    tmp = MsgBox("上の色と同じにしますか？（下の場合は、いいえ）", vbYesNoCancel)
    If tmp = vbCancel Then Exit Sub
    
    If Selection.Count > 1 Then
        For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
            If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上の値のみ取り出す
                If tmp = vbYes Then
                    i = -1
                    
                    Do
                        If myRange.Offset(i, 0).EntireRow.Hidden = False Then
                            myRange.Interior.Color = myRange.Offset(i, 0).DisplayFormat.Interior.Color
                            Exit Do
                            
                        End If
                        
                        i = i - 1
                        
                    Loop
                    
                ElseIf tmp = vbNo Then
                    i = 1
                    
                    Do
                        If myRange.Offset(i, 0).EntireRow.Hidden = False Then
                            myRange.Interior.Color = myRange.Offset(i, 0).DisplayFormat.Interior.Color
                            Exit Do
                            
                        End If
                        
                        i = i + 1
                        
                    Loop
                    
                End If
                
            End If
            
        Next myRange
        
    Else
        If tmp = vbYes Then
            i = -1
            
            Do
                If ActiveCell.Offset(i, 0).EntireRow.Hidden = False Then
                    ActiveCell.Interior.Color = ActiveCell.Offset(i, 0).DisplayFormat.Interior.Color
                    Exit Do
                    
                End If
                
                i = i - 1
                
            Loop
            
        ElseIf tmp = vbNo Then
            i = 1
            
            Do
                If ActiveCell.Offset(i, 0).EntireRow.Hidden = False Then
                    ActiveCell.Interior.Color = ActiveCell.Offset(i, 0).DisplayFormat.Interior.Color
                    Exit Do
                    
                End If
                
                i = i + 1
                
            Loop
            
        End If
        
    End If
    
End Sub