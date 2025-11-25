Sub 空白改行区切りでクリップボードに格納()
    'Microsoft Forms 2.0 Object Libraryを参照設定して下さい
    '矩形範囲を選択する
    '列を空白で区切り、行を改行で区切ります
    Dim i As Long
    Dim j As Long

    For i = Selection.Row To Selection.Rows(Selection.Rows.Count).Row
        For j = Selection.Column To Selection.Columns(Selection.Columns.Count).Column
            If Cells(i, j).Address = Cells(i, j).MergeArea(1).Address And _
            Rows(i).Hidden = False And Columns(j).Hidden = False Then '結合セルの場合は左上の値のみ取り出す。非表示セルは処理しない
                If Cells(i, j).MergeArea(1).Address = Cells(i, Selection. _
                Columns(Selection.Columns.Count).Column).MergeArea(1).Address Then
                '選択範囲の最終列（最終列を含む結合セル）であれば末尾に改行を追加
                    If InStr(Cells(i, j), vbLf) = 0 Then
                        V = V & Cells(i, j).Value & vbCrLf
                    Else
                        V = V & """" & Cells(i, j).Value & """" & vbCrLf    'セル内改行があれば前後を""で囲む
                    End If
                Else
                '選択範囲の最終列（最終列を含む結合セル）以外は末尾にTabを追加
                    If InStr(Cells(i, j), vbLf) = 0 Then
                        V = V & Cells(i, j).Value
                    Else
                        V = V & """" & Cells(i, j).Value & """"   'セル内改行があれば前後を""で囲む
                    End If
                End If
            End If
        Next j
    Next i
    
    V = Left(V, Len(V) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）
    
    With New MSForms.DataObject
        .SetText V  '変数の値をDataObjectに格納する
        .PutInClipboard 'DataObjectのデータをクリップボードに格納する
    End With
    
End Sub