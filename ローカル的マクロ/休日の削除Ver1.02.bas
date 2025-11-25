Sub 休日の削除()
    'ローカル的なマクロです
    '重複時の削除のみで復活はしません
    Dim 列始 As Integer
    Dim 列終 As Integer
    Dim 日付列 As Long
    Dim 内容列 As Long
    Dim 行終 As Long
    Dim 列 As Integer
    Dim 行 As Long
    Dim 日付 As String
    Dim i As Long
    Dim 削除 As Long
    
    列始 = Cells.Find("A", LookAt:=xlWhole).Column
    列終 = Cells.Find("B", LookAt:=xlWhole).Column
    日付列 = 1
    内容列 = 6
    
    行終 = Cells(Rows.Count, 日付列).End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    For 列 = 列始 To 列終
        For 行 = 1 To 行終
            If Cells(行, 内容列).Value Like "指定休*休暇*" And TypeName(Cells(行, 日付列).Value) = "Date" Then
                日付 = Year(Cells(行, 日付列).Value) & Month(Cells(行, 日付列).Value) & Day(Cells(行, 日付列).Value)
                
                For i = 行 + 1 To 行終
                    If TypeName(Cells(i, 日付列).Value) = "Date" Then
                        If 日付 = Year(Cells(i, 日付列).Value) & Month(Cells(i, 日付列).Value) & Day(Cells(i, 日付列).Value) Then
                            If Cells(i, 列).Value <> "" Then
                                If Cells(行, 列).Value <> "" And Not Cells(行, 列).Value Like "*[pPｐＰaAａＡfFｆＦ]*" Then
                                    Cells(行, 列).ClearContents
                                    削除 = 削除 + 1
                                End If
                                
                                Exit For
                                
                            End If
                        Else
                            Exit For
                            
                        End If
                    End If
                Next i

            End If

        Next 行
        
    Next 列
    
    Application.ScreenUpdating = True

    MsgBox "削除：" & 削除

End Sub