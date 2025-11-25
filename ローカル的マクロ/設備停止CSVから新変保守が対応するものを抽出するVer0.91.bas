Sub 設備停止CSVから新変保守が対応するものを抽出する()
    'ローカル的なマクロです
    Dim 電気所 As Variant
    Dim 線路名 As Variant
    Dim Keyword1 As String
    Dim 列要求 As Long
    Dim 列実施 As Long
    Dim 列一 As Long
    Dim 列二 As Long
    Dim 列三 As Long
    Dim 列四 As Long
    Dim 列五 As Long
    Dim 列六 As Long
    Dim 列七 As Long
    Dim 列八 As Long
    Dim 列九 As Long
    Dim 列接地 As Long
    Dim 行始 As Long
    Dim 行終 As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    電気所 = Array("あ", "い", "う", "え", "お")    '必要に応じて変更する
    線路名 = Array("入善北線") '必要に応じて変更する
    Keyword1 = "東魚津支線" '必要に応じて変更する
    
    '---ここから行番号・列番号を取得する処理---
    
    列要求 = Cells.Find("要求箇所", LookAt:=xlWhole).Column
    列実施 = Cells.Find("実施箇所", LookAt:=xlWhole).Column
    列一 = Cells.Find("停止設備および線路名　停止設備１", LookAt:=xlWhole).Column
    列二 = Cells.Find("停止設備および線路名　停止設備２", LookAt:=xlWhole).Column
    列三 = Cells.Find("停止設備および線路名　停止設備３", LookAt:=xlWhole).Column
    列四 = Cells.Find("停止設備および線路名　停止設備４", LookAt:=xlWhole).Column
    列五 = Cells.Find("停止設備および線路名　停止設備５", LookAt:=xlWhole).Column
    列六 = Cells.Find("停止設備および線路名　停止設備６", LookAt:=xlWhole).Column
    列七 = Cells.Find("停止設備および線路名　停止設備７", LookAt:=xlWhole).Column
    列八 = Cells.Find("停止設備および線路名　停止設備８", LookAt:=xlWhole).Column
    列九 = Cells.Find("停止設備および線路名　停止設備９", LookAt:=xlWhole).Column
    列接地 = Cells.Find("給電接地　接地有無", LookAt:=xlWhole).Column
    
    行始 = Cells.Find("要求箇所", LookAt:=xlWhole).Row
    行終 = Cells(Rows.Count, 列要求).End(xlUp).Row
    
    '---ここから絞り込みをする処理---
    
    For i = 行始 To 行終
    
        If Cells(i, 列要求).Value Like "新[変電]保守" Then   '要求箇所での絞り込み
            Cells(i, "A").Interior.Color = 15921906
            GoTo skip
            
        End If
    
        j = 0
    
        Do While j <= UBound(電気所)
            If Cells(i, 列実施).Value = 電気所(j) Then   '実施箇所での絞り込み
                Cells(i, "A").Interior.Color = 15921906
                GoTo skip
                
            Else
                j = j + 1
            End If
        Loop
        
        k = 0
    
        Do While k <= UBound(線路名)    '22kV線路停止かつ給電接地有りで絞り込み
            If (Cells(i, 列一) = 線路名(k) Or Cells(i, 列二) = 線路名(k) Or Cells(i, 列三) = 線路名(k) _
            Or Cells(i, 列四) = 線路名(k) Or Cells(i, 列五) = 線路名(k) Or Cells(i, 列六) = 線路名(k) _
            Or Cells(i, 列七) = 線路名(k) Or Cells(i, 列八) = 線路名(k) Or Cells(i, 列九) = 線路名(k)) _
            And Cells(i, 列接地) = "1" Then
                Cells(i, "A").Interior.Color = 15921906
                GoTo skip
                
            Else
                k = k + 1
            End If
        Loop
        
        'Keyword1で絞り込み
        If Cells(i, 列一) = Keyword1 Or Cells(i, 列二) = Keyword1 Or Cells(i, 列三) = Keyword1 _
        Or Cells(i, 列四) = Keyword1 Or Cells(i, 列五) = Keyword1 Or Cells(i, 列六) = Keyword1 _
        Or Cells(i, 列七) = Keyword1 Or Cells(i, 列八) = Keyword1 Or Cells(i, 列九) = Keyword1 Then
            Cells(i, "A").Interior.Color = 15921906
            GoTo skip
            
        End If
        
skip:
        
    Next i
    
    Cells(行始, "A").AutoFilter Field:=1  '当該セルを含むアクティブセル領域をオートフィルターに設定する。引数は既にオートフィルターがある場合に解除しないためのダミー
    
    If ActiveSheet.FilterMode = True Then ActiveSheet.ShowAllData   'オートフィルターによる絞り込みがあれば一旦すべてクリアする
    
    Cells(行始, "A").AutoFilter Field:=1, Criteria1:=15921906, Operator:=xlFilterCellColor    'セルの背景色で絞り込みを行う
    
End Sub