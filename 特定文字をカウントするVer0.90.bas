Sub 特定文字をカウントする()
    '選択範囲に対して処理を行います
    '全角・半角および大文字・小文字を区別しません
    Dim 特定文字 As String
    Dim tmp As String
    Dim myRange As Range
    Dim 配列 As Variant
    Dim cnt As Integer
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    特定文字 = InputBox("カウントする文字を入力して下さい。")
    If 特定文字 = "" Then Exit Sub
    
    tmp = StrConv(StrConv(特定文字, vbLowerCase), vbNarrow)    '半角小文字に統一
    
    If Selection.Count > 1 Then
        For Each myRange In Selection.SpecialCells(xlCellTypeVisible)   '可視セルのみに処理を行う
            If myRange.Value <> "" Then   '空白以外のセルに処理を行う（理由：空白のみをSplitすると要素とデータのない配列を返し，UBoundは-1を返すため）
                配列 = Split(StrConv(StrConv(myRange.Value, vbLowerCase), vbNarrow), tmp)  '半角小文字に統一したセル内の文字をtmpで分割し配列に格納する
                cnt = cnt + UBound(配列)
                
            End If
            
        Next myRange
        
    Else
        If ActiveCell.Value <> "" Then
            配列 = Split(StrConv(StrConv(ActiveCell.Value, vbLowerCase), vbNarrow), tmp)
            cnt = UBound(配列)
            
        End If
        
    End If
    
    myLib.SetText cnt  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する
    
    MsgBox 特定文字 & " は " & cnt & " 個ありました。" & vbCrLf & "（ " & cnt & " をクリップボードに転送しました）"
    
End Sub