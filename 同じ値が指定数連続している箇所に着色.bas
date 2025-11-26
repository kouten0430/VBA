Sub 同じ値が指定数連続している箇所に着色()
    Dim 値 As Variant
    Dim 連続数 As Long
    Dim 色 As Long
    Dim 行 As Long
    Dim 行終 As Long
    Dim 列 As Long
    Dim 列終 As Long
    Dim myUni As Range

    Application.DisplayAlerts = False
    
    値 = ""
    連続数 = 7
    色 = 255

    行 = Selection.Row + 1
    行終 = Selection.Rows(Selection.Rows.Count).Row
    列 = Selection.Column
    列終 = Selection.Columns(Selection.Columns.Count).Column
    
    For i = 列 To 列終
        For j = 行 To 行終
            If Cells(j, i).Value = 値 And Cells(j - 1, i).Value = Cells(j, i).Value Then
                If myUni Is Nothing Then
                    Set myUni = Range(Cells(j - 1, i), Cells(j, i))
                Else
                    Set myUni = Union(myUni, Cells(j, i))
                End If
            Else
                If Not myUni Is Nothing Then
                    If myUni.Count >= 連続数 Then
                        myUni.Interior.Color = 色
                    End If
                    Set myUni = Nothing
                End If
            End If
    
        Next j
        
        If Not myUni Is Nothing Then
            If myUni.Count >= 連続数 Then
                myUni.Interior.Color = 色
            End If
            Set myUni = Nothing
        End If

    Next i

End Sub