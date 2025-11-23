Sub ビットマップ画像をセルで描画する()
    '24ビットビットマップのみ処理可能
    Dim ファイル名 As Variant
    Dim 配列() As Byte
    Dim i As Long
    Dim tmp As String
    Dim 横 As Long
    Dim 縦 As Long
    Dim 詰物 As Integer
    Dim 行 As Long
    Dim 列 As Long
    Dim 色 As String
    
    ファイル名 = Application.GetOpenFilename(Title:="ビットマップ画像を選択")
        If TypeName(ファイル名) = "Boolean" Then Exit Sub
    
    Open ファイル名 For Binary As #1
        ReDim 配列(LOF(1))
        Get #1, , 配列
    Close #1
    
    tmp = ""
    For i = 18 To 21    '画像の横幅（ピクセル）を取得
        tmp = WorksheetFunction.Dec2Hex(配列(i), 2) & tmp
    Next i
    横 = WorksheetFunction.Hex2Dec(tmp)
    
    tmp = ""
    For i = 22 To 25    '画像の縦幅（ピクセル）を取得
        tmp = WorksheetFunction.Dec2Hex(配列(i), 2) & tmp
    Next i
    縦 = WorksheetFunction.Hex2Dec(tmp)
    
    If (横 * 3 Mod 4) <> 0 Then 詰物 = 4 - (横 * 3 Mod 4)   '横幅が4バイトで割り切れない場合の詰物の数を算出
    
    i = 54   '画像データ開始位置
    
    For 行 = 縦 To 1 Step -1
        For 列 = 1 To 横
            色 = WorksheetFunction.Dec2Hex(配列(i), 2) _
            & WorksheetFunction.Dec2Hex(配列(i + 1), 2) _
            & WorksheetFunction.Dec2Hex(配列(i + 2), 2)
            Cells(行, 列).Interior.Color = WorksheetFunction.Hex2Dec(色)
            i = i + 3
            DoEvents
        Next 列
        
        i = i + 詰物
        
    Next 行
End Sub