Sub 水平方向に離れたセルのデータをCBに格納する列固定版()
    '選択したセルから水平方向に離れたセルのデータを改行区切りでCBに格納します
    '選択したセルを青色に塗りつぶします
    '内容が複数の場合は、水色→青色の順に塗りつぶします
    '列全体を網かけします
    '月間予定の日付をもとに週間予定をスクロールさせます
    'ローカル的なマクロです
    Dim 列 As Long
    Dim 複数の区切り As String
    Dim myRange As Range
    Dim V As String
    Dim myLib As Object
    Set myLib = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")  '参照設定なしでDataObjectのインスタンスを生成する
    
    '---ここから転記先をスクロールさせる処理---

    Dim 転記先シート As Worksheet
    Dim 日付列 As Long
    Dim 行終 As Long
    Dim 今日 As String
    Dim i As Long
    Dim tmp As String
    
    Set 転記先シート = Workbooks(2).ActiveSheet
    日付列 = 1
    行終 = 9999
    
    For i = 1 To Selection(1).Row
        If TypeName(Cells(i, Selection(1).Column).Value) = "Date" Then
            今日 = Day(Cells(i, Selection(1).Column).Value)    '月間予定から日付を取得
            
            Exit For
            
        End If
        
    Next i
    
    For i = 1 To 行終
        If TypeName(転記先シート.Cells(i, 日付列).Value) = "Date" Then
            tmp = Day(転記先シート.Cells(i, 日付列).Value)
            
            If 今日 = tmp Then
                Windows(2).ScrollRow = i
                Windows(2).ScrollColumn = 日付列
                
                Exit For
                
            End If
            
        End If
        
    Next i
    
    '---ここまで---
    
    列 = 2
    複数の区切り = vbLf & "--------" & vbLf
    
    For Each myRange In Selection
        V = V & Cells(myRange.Row, 列).Value & vbCrLf
        
        If InStr(myRange.Value, 複数の区切り) <> 0 Then   '内容が複数の場合の処理
            If myRange.Interior.Color <> 15986394 Then
                myRange.Interior.Color = 15986394   '一回目は水色に塗りつぶす
            Else
                myRange.Interior.Color = 15773696   '二回目は青色に塗りつぶす
            End If
        Else    '内容が単数の場合の処理
            myRange.Interior.Color = 15773696
        End If

        myRange.EntireColumn.Interior.Pattern = xlPatternGray8
        
    Next myRange
        
    V = Left(V, Len(V) - 2) '最終行の改行区切りを取り除く（CrLfは2文字）
    
    myLib.SetText V  '変数の値をDataObjectに格納する
    myLib.PutInClipboard 'DataObjectのデータをクリップボードに格納する

End Sub