Sub 背景色および文字色の16進数を取得()
    'CSS用にアクティブセル背景色および文字色から16進数を取得します
    'Hex関数で取得した値はBGRの順なので並び替えが必要（CSSはRGBの順）
    Dim mb As Integer
    Dim cc As Long
    Dim hx As String
    Dim bl As String
    Dim gr As String
    Dim re As String
    Dim V As String
    Dim blue As Integer
    Dim green As Integer
    Dim red As Integer
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
    
    mb = MsgBox("OK：背景色を取得" & vbCrLf & "ｷｬﾝｾﾙ：文字色を取得", vbOKCancel)
        If mb = 1 Then
            cc = ActiveCell.Interior.Color
        Else
            cc = ActiveCell.Font.Color
        End If
        
    hx = Right("00000" & Hex(cc), 6)
    bl = Mid(hx, 1, 2)
    gr = Mid(hx, 3, 2)
    re = Mid(hx, 5, 2)
    V = re & gr & bl
    blue = CInt("&H" & bl)
    green = CInt("&H" & gr)
    red = CInt("&H" & re)
    MsgBox "RGB値:" & red & "," & green & "," & blue & vbCrLf & "16進数:" & V
    
    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する
    MsgBox "16進数:" & V & vbCrLf & "をｸﾘｯﾌﾟﾎﾞｰﾄﾞに転送しました！"
End Sub