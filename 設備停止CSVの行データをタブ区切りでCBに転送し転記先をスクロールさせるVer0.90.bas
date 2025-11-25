Sub 設備停止CSVの行データをタブ区切りでCBに転送し転記先をスクロールさせる()
    '事前準備として，作業票側件名Noを月間側にコピペし重複しない値を調べておく
    '処理が必要な行を選択して実行する
    '転記先（月間側）で貼り付けする
    Dim 行 As Long
    Dim 件名No As String
    Dim myRange As Range
    Dim CB As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    行 = Selection.Row
    件名No = Cells(行, 1).Value
    
    For Each myRange In Range(Cells(行, "A"), Cells(行, "FK"))
        CB = CB & myRange.Value & vbTab
    
    Next myRange
    
    CB = Left(CB, Len(CB) - 1) '最終行のタブ区切りを取り除く
    
    myLib.SetText CB  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する

    '---ここから転記先をスクロールさせる処理---

    Dim 転記先シート As Worksheet
    Dim i As Long
    Dim 行終 As Long
    Dim j As Long
    
    If Workbooks.Count = 3 Then
        For i = 1 To Workbooks.Count    '専用のプロパティがないため、ループでアクティブブックのインデックスを調べる
            If Workbooks(i).Name = ActiveWorkbook.Name Then
                If i = 2 Then
                    Set 転記先シート = Workbooks(i + 1).ActiveSheet
                    Exit For
                    
                ElseIf i = 3 Then
                    Set 転記先シート = Workbooks(i - 1).ActiveSheet
                    Exit For
    
                End If
                
            End If
            
        Next i
        
    Else
        MsgBox "ブックを2つだけ開いた状態にして下さい。"
        Exit Sub
        
    End If
    
    行終 = 転記先シート.Cells(Rows.Count, 1).End(xlUp).Row
    
    For j = 1 To 行終
        If 転記先シート.Cells(j, 1).Value = 件名No Then
            Windows(2).ScrollRow = j
            Windows(2).ScrollColumn = 1
            Windows(2).Activate
            
            Exit For
            
        End If
        
    Next j
    
    '---ここまで---
    
End Sub