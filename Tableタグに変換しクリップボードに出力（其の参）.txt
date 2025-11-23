Sub 選択範囲をTableタグに変換しクリップボードに出力其の参()
    '正方形または長方形のような連続した選択範囲とする
    'セルの内容の水平位置を再現します（左、右、中央のみ）。水平位置が無指定なら無指定であることを再現します。
    Dim i As Long
    Dim j As Long
    Dim V As String
    Dim rh As Integer
    Dim ch As Integer
    Dim ha As Integer
    Dim Alg As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    rh = MsgBox("選択範囲の上端の行を" & vbCrLf & "見出しにしますか？", vbYesNo)
    ch = MsgBox("選択範囲の左端の列を" & vbCrLf & "見出しにしますか？", vbYesNo)
    
    V = "<table>" & vbCrLf

    For i = Selection.Row To Selection.Rows(Selection.Rows.Count).Row
        For j = Selection.Column To Selection.Columns(Selection.Columns.Count).Column
            If i = Selection.Row And rh = 6 Then '見出し行の処理。データを<th></th>で囲む
                If j = Selection.Column Then    '選択範囲の左端であれば冒頭に<tr>を追加
                    ha = Cells(i, j).HorizontalAlignment    'データの水平位置を取得する
                        If ha = -4131 Then
                            Alg = " align=""left"""
                        ElseIf ha = -4152 Then
                            Alg = " align=""right"""
                        ElseIf ha = -4108 Then
                            Alg = " align=""center"""
                        Else
                            Alg = ""
                        End If
                    V = V & "<tr>" & vbCrLf & "<th" & Alg & ">" & Cells(i, j).Value & "</th>"
                        If j = Selection.Columns(Selection.Columns.Count).Column Then     '左端かつ右端である場合の処置
                            V = V & vbCrLf & "</tr>" & vbCrLf
                        End If
                ElseIf j = Selection.Columns(Selection.Columns.Count).Column Then   '選択範囲の右端であれば末尾に</tr>を追加
                    ha = Cells(i, j).HorizontalAlignment    'データの水平位置を取得する
                        If ha = -4131 Then
                            Alg = " align=""left"""
                        ElseIf ha = -4152 Then
                            Alg = " align=""right"""
                        ElseIf ha = -4108 Then
                            Alg = " align=""center"""
                        Else
                            Alg = ""
                        End If
                    V = V & "<th" & Alg & ">" & Cells(i, j).Value & "</th>" & vbCrLf & "</tr>" & vbCrLf
                Else    '左端と右端以外の処理
                    ha = Cells(i, j).HorizontalAlignment    'データの水平位置を取得する
                        If ha = -4131 Then
                            Alg = " align=""left"""
                        ElseIf ha = -4152 Then
                            Alg = " align=""right"""
                        ElseIf ha = -4108 Then
                            Alg = " align=""center"""
                        Else
                            Alg = ""
                        End If
                    V = V & "<th" & Alg & ">" & Cells(i, j).Value & "</th>"
                End If
            Else '見出し行以外の処理
                If j = Selection.Column Then    '選択範囲の左端であれば冒頭に<tr>を追加
                    If ch = 6 Then  '見出し列の処理。データを<th></th>で囲む
                        ha = Cells(i, j).HorizontalAlignment    'データの水平位置を取得する
                            If ha = -4131 Then
                                Alg = " align=""left"""
                            ElseIf ha = -4152 Then
                                Alg = " align=""right"""
                            ElseIf ha = -4108 Then
                                Alg = " align=""center"""
                            Else
                                Alg = ""
                            End If
                        V = V & "<tr>" & vbCrLf & "<th" & Alg & ">" & Cells(i, j).Value & "</th>"
                    Else    '見出し列以外の処理。データを<td></td>で囲む
                        ha = Cells(i, j).HorizontalAlignment    'データの水平位置を取得する
                            If ha = -4131 Then
                                Alg = " align=""left"""
                            ElseIf ha = -4152 Then
                                Alg = " align=""right"""
                            ElseIf ha = -4108 Then
                                Alg = " align=""center"""
                            Else
                                Alg = ""
                            End If
                        V = V & "<tr>" & vbCrLf & "<td" & Alg & ">" & Cells(i, j).Value & "</td>"
                    End If
                        If j = Selection.Columns(Selection.Columns.Count).Column Then     '左端かつ右端である場合の処置
                            V = V & vbCrLf & "</tr>" & vbCrLf
                        End If
                ElseIf j = Selection.Columns(Selection.Columns.Count).Column Then   '選択範囲の右端であれば末尾に</tr>を追加
                    ha = Cells(i, j).HorizontalAlignment    'データの水平位置を取得する
                        If ha = -4131 Then
                            Alg = " align=""left"""
                        ElseIf ha = -4152 Then
                            Alg = " align=""right"""
                        ElseIf ha = -4108 Then
                            Alg = " align=""center"""
                        Else
                            Alg = ""
                        End If
                    V = V & "<td" & Alg & ">" & Cells(i, j).Value & "</td>" & vbCrLf & "</tr>" & vbCrLf
                Else    '左端と右端以外の処理
                    ha = Cells(i, j).HorizontalAlignment    'データの水平位置を取得する
                        If ha = -4131 Then
                            Alg = " align=""left"""
                        ElseIf ha = -4152 Then
                            Alg = " align=""right"""
                        ElseIf ha = -4108 Then
                            Alg = " align=""center"""
                        Else
                            Alg = ""
                        End If
                    V = V & "<td" & Alg & ">" & Cells(i, j).Value & "</td>"
                End If
            End If
        Next j
    Next i
    
    V = V & "</table>"
    
    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する
    
    MsgBox "ＨＴＭＬをクリップボードに" & vbCrLf & "出力しました！" & vbCrLf & vbCrLf & _
    "ブログなどでお好みの位置にペースト" & vbCrLf & "して下さい。"
    
End Sub