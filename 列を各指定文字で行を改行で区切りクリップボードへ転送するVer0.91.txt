Sub 列を各指定文字で行を改行で区切りクリップボードへ転送する()
    '選択された矩形範囲のデータを、列を各指定文字で、行を改行で区切りクリップボードへ転送します
    '行の冒頭、末尾にも文字列を入れることができます。不要ならInputBoxで空白を指定して下さい。
    Dim y As Long
    Dim x As Long
    Dim Tdim() As String
    Dim dim1 As Integer
    Dim dim2 As Integer
    Dim j As Long
    Dim i As Long
    Dim myRange As Range
    Dim dc() As String
    Dim V As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    If Selection.Areas.Count > 1 Then   '複数の矩形範囲が選択されている場合は終了する
        MsgBox "一つの矩形範囲のみ選択して再度実行して下さい。"
        Exit Sub
    End If
    
    
    '---ここから可視セル範囲を二次元配列に格納する処理---
    
    y = Selection.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count
    x = Selection.Rows(1).SpecialCells(xlCellTypeVisible).Cells.Count
    
    ReDim Tdim(1 To y, 1 To x)
    
    dim1 = 1
    dim2 = 1
        
    For i = Selection.Row To Selection.Rows(Selection.Rows.Count).Row
        For j = Selection.Column To Selection.Columns(Selection.Columns.Count).Column
            If Rows(i).Hidden = False And Columns(j).Hidden = False Then '非表示セルは処理しない
                Tdim(dim1, dim2) = Cells(i, j).Value
                    dim2 = dim2 + 1
                    If dim2 > x Then
                        dim2 = 1
                        dim1 = dim1 + 1
                    End If
            End If
        Next j
    Next i
    
    '---ここまで---
    
    
    ReDim dc(1 To UBound(Tdim, 2) + 1)  '+1は末尾に挿入する文字列を格納するため
    
    For j = 1 To UBound(Tdim, 2) + 1
        If j < UBound(Tdim, 2) + 1 Then
            dc(j) = Application.InputBox(prompt:=j & "列目の前に挿入する文字列", Type:=2)
                If dc(j) = "False" Then
                    Exit Sub
                End If
        Else
            dc(j) = Application.InputBox(prompt:="末尾に挿入する文字列", Type:=2)
                If dc(j) = "False" Then
                    Exit Sub
                End If
        End If
    Next j

    For i = 1 To UBound(Tdim, 1)
        For j = 1 To UBound(Tdim, 2)
            If j < UBound(Tdim, 2) Then '選択範囲の最終列以外の処理
                V = V & dc(j) & Tdim(i, j)
            Else    '選択範囲の最終列の処理
                V = V & dc(j) & Tdim(i, j) & dc(j + 1) & vbCrLf
            End If
        Next j
    Next i
    
    V = Left(V, Len(V) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）

    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する
    
End Sub