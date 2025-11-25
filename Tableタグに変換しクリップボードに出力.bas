Sub 選択範囲をTableタグに変換しクリップボードに出力()
    'Microsoft Forms 2.0 Object Libraryを参照設定して下さい
    '正方形または長方形のような連続した選択範囲とする
    Dim i As Long
    Dim j As Long
    Dim V As String
    Dim rh As Integer
    Dim ch As Integer
    Dim myLib As New DataObject
    
    rh = MsgBox("選択範囲の上端の行を" & vbCrLf & "見出しにしますか？", vbYesNo)
    ch = MsgBox("選択範囲の左端の列を" & vbCrLf & "見出しにしますか？", vbYesNo)
    
    V = "<table>" & vbCrLf

    For i = Selection.Row To Selection.Rows(Selection.Rows.Count).Row
        For j = Selection.Column To Selection.Columns(Selection.Columns.Count).Column
            If i = Selection.Row And rh = 6 Then '見出し行の処理。データを<th></th>で囲む
                If j = Selection.Column Then    '選択範囲の左端であれば冒頭に<tr>を追加
                    V = V & "<tr>" & vbCrLf & "<th>" & Cells(i, j).Value & "</th>"
                ElseIf j = Selection.Columns(Selection.Columns.Count).Column Then   '選択範囲の右端であれば末尾に</tr>を追加
                    V = V & "<th>" & Cells(i, j).Value & "</th>" & vbCrLf & "</tr>" & vbCrLf
                Else    '左端と右端以外の処理
                    V = V & "<th>" & Cells(i, j).Value & "</th>"
                End If
            Else '見出し行以外の処理。データを<td></td>で囲む
                If j = Selection.Column Then    '選択範囲の左端であれば冒頭に<tr>を追加
                    If ch = 6 Then  '見出し列の処理。データを<th></th>で囲む
                        V = V & "<tr>" & vbCrLf & "<th>" & Cells(i, j).Value & "</th>"
                    Else    '見出し列以外の処理。データを<td></td>で囲む
                        V = V & "<tr>" & vbCrLf & "<td>" & Cells(i, j).Value & "</td>"
                    End If
                ElseIf j = Selection.Columns(Selection.Columns.Count).Column Then   '選択範囲の右端であれば末尾に</tr>を追加
                    V = V & "<td>" & Cells(i, j).Value & "</td>" & vbCrLf & "</tr>" & vbCrLf
                Else    '左端と右端以外の処理
                    V = V & "<td>" & Cells(i, j).Value & "</td>"
                End If
            End If
        Next j
    Next i
    
    V = V & "</table>"
    
    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する
    
    MsgBox "Tableタグ(HTML)をクリップボードに" & vbCrLf & "出力しました！" & vbCrLf & vbCrLf & _
    "ブログなどでお好みの位置にペースト" & vbCrLf & "して下さい。"
    
End Sub