Sub Tab改行区切りでクリップボードに格納()
    Dim i As Long
    Dim j As Long
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    If Selection.Areas.Count > 1 Then   '複数の矩形範囲が選択されている場合は終了する
        MsgBox "一つの矩形範囲のみ選択して再度実行して下さい。"
        Exit Sub
    End If

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
                        V = V & Cells(i, j).Value & vbTab
                    Else
                        V = V & """" & Cells(i, j).Value & """" & vbTab   'セル内改行があれば前後を""で囲む
                    End If
                End If
            End If
        Next j
    Next i
    
    V = Left(V, Len(V) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）
    
    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する
    
End Sub