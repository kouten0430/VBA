Sub 保全計画と設備停止計画のチェックBPD()
    '検索元と検索先のブックを開いておく
    '検索元のブックをアクティブにし、カレントリージョンが効く位置でマクロを実行する
    'Workbooksからのオブジェクト指定が無いものは保全計画側の処理です
    'クラスおよび設備停止有無でフィルタをかけてから実行するとより確実です（可視セルのみ処理対象のため）
    Dim tmp As Variant
    Dim 検索先ワークブック名 As String
    Dim 検索元ワークブック名 As String
    Dim 電気所等名称列 As Integer
    Dim 保全計画テキスト列 As Integer
    Dim クラス列 As Integer
    Dim 月範囲 As Range
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 列始 As Integer
    Dim 列終 As Integer
    Dim 年停行始 As Long
    Dim 年停行終 As Long
    Dim 作業内容列 As Integer
    Dim 実施箇所列 As Integer
    Dim 開始日時列 As Integer
    Dim i As Long
    Dim 電気所等名称 As String
    Dim 保全計画テキスト As String
    Dim 切始 As Integer
    Dim 切出し文字 As String
    Dim 作業内容 As String
    Dim 実施個所 As String
    Dim 月 As Integer
    Dim 要求時注釈列 As Integer
    Dim 要求時注釈 As String
    
    機器種類 = "ＰＤ"

    検索先ワークブック名 = InputBox("設備停止側のブック名を拡張子を含めて入力して下さい。")
    検索元ワークブック名 = ActiveWorkbook.Name
    電気所等名称列 = 2  '必要に応じて変更
    保全計画テキスト列 = 6  '必要に応じて変更
    クラス列 = 32   '必要に応じて変更
    Set 月範囲 = Range("M3:X3") '必要に応じて変更

    行始 = ActiveCell.CurrentRegion.Row
    行終 = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row
    列始 = ActiveCell.CurrentRegion.Column
    列終 = ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column
    年停行始 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Row  'Rangeの引数はお好みで
    年停行終 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows(Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows.Count).Row 'Rangeの引数はお好みで
    作業内容列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("作業内容", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    実施箇所列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("実施箇所", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    開始日時列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求期間　開始日時", LookAt:=xlWhole).Column   'Rangeの引数は必要に応じて変更
    要求時注釈列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求時注釈", LookAt:=xlWhole).Column
    
    For i = 行始 To 行終
    
        If Rows(i).Hidden = False And Cells(i, 電気所等名称列).Value Like "*" & "　" & "*" Then  '可視セルかつ電気所名に全角スペースが含まれる場合のみ処理を行う

            電気所等名称 = Left(Cells(i, 電気所等名称列).Value, InStr(Cells(i, 電気所等名称列).Value, "　") - 1)    '左端から電気所名のみを切り出す
            If Right(電気所等名称, 1) = "開" Then
                電気所等名称 = 電気所等名称 & "閉所"
            Else
                電気所等名称 = 電気所等名称 & "電所"
            End If
            
            保全計画テキスト = StrConv(Cells(i, 保全計画テキスト列).Value, vbWide)    '全角にする
            保全計画テキスト = StrConv(保全計画テキスト, vbUpperCase) '大文字にする
            保全計画テキスト = Replace(保全計画テキスト, "　ＢＰＤ", "ＢＰＤ")
    
            If 保全計画テキスト Like "*" & 機器種類 & "*" And (Cells(i, クラス列).Value = "X_APT_") And Not 保全計画テキスト Like "*" & "ＬＰＤ" & "*" Then
    
                切始 = InStr(保全計画テキスト, 機器種類)
                
                For j = 切始 - 1 To 1 Step -1
                    If Mid(保全計画テキスト, j, 1) Like "[!１-３ＡＢ母線]" Then
                        切出し文字 = Mid(保全計画テキスト, j + 1, 切始 + Len(機器種類) - 1 - j)
                        切出し文字 = Replace(切出し文字, "母線", "Ｂ")
                        Exit For
                        
                    End If
                Next j

                For k = 年停行始 To 年停行終
                
                    If Workbooks(検索先ワークブック名).Sheets(1).Rows(k).Hidden = False Then  '可視セルのみ処理を行う
                
                        作業内容 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Value, vbWide)  '全角にする
                        作業内容 = StrConv(作業内容, vbUpperCase) '大文字にする
                        作業内容 = Replace(作業内容, "ブス", "Ｂ")
                        作業内容 = Replace(作業内容, "母線", "Ｂ")
                        
                        要求時注釈 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Value, vbWide)  '全角にする
                        要求時注釈 = StrConv(要求時注釈, vbUpperCase) '大文字にする
                        要求時注釈 = Replace(要求時注釈, "ブス", "Ｂ")
                        要求時注釈 = Replace(要求時注釈, "母線", "Ｂ")
                        
                        実施個所 = Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 実施箇所列).Value
                
                        If 実施個所 = 電気所等名称 And 作業内容 Like "*" & 切出し文字 & "*" Then    '実施個所と作業内容が一致していれば着色
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                            
                        ElseIf 実施個所 = 電気所等名称 And 要求時注釈 Like "*" & 切出し文字 & "*" Then
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                        
                        End If
                    
                    End If
                    
                    DoEvents
                
                Next k

            End If

        End If

        Application.StatusBar = "処理実行中．．．" & Round(i / 行終 * 100, 0) & "%"
        
    Next i
    
    Application.StatusBar = False
    
End Sub
Sub 保全計画と設備停止計画のチェックLPD()
    '検索元と検索先のブックを開いておく
    '検索元のブックをアクティブにし、カレントリージョンが効く位置でマクロを実行する
    'Workbooksからのオブジェクト指定が無いものは保全計画側の処理です
    'クラスおよび設備停止有無でフィルタをかけてから実行するとより確実です（可視セルのみ処理対象のため）
    Dim tmp As Variant
    Dim 検索先ワークブック名 As String
    Dim 検索元ワークブック名 As String
    Dim 電気所等名称列 As Integer
    Dim 保全計画テキスト列 As Integer
    Dim クラス列 As Integer
    Dim 月範囲 As Range
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 列始 As Integer
    Dim 列終 As Integer
    Dim 年停行始 As Long
    Dim 年停行終 As Long
    Dim 作業内容列 As Integer
    Dim 実施箇所列 As Integer
    Dim 開始日時列 As Integer
    Dim i As Long
    Dim 電気所等名称 As String
    Dim 保全計画テキスト As String
    Dim 切始 As Integer
    Dim 切出し文字 As String
    Dim 作業内容 As String
    Dim 実施個所 As String
    Dim 月 As Integer
    Dim 要求時注釈列 As Integer
    Dim 要求時注釈 As String
    Dim 停止設備列 As Integer
    Dim 停止設備 As String
    
    機器種類 = "ＬＰＤ"

    検索先ワークブック名 = InputBox("設備停止側のブック名を拡張子を含めて入力して下さい。")
    検索元ワークブック名 = ActiveWorkbook.Name
    電気所等名称列 = 2  '必要に応じて変更
    保全計画テキスト列 = 6  '必要に応じて変更
    クラス列 = 32   '必要に応じて変更
    Set 月範囲 = Range("M3:X3") '必要に応じて変更

    行始 = ActiveCell.CurrentRegion.Row
    行終 = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row
    列始 = ActiveCell.CurrentRegion.Column
    列終 = ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column
    年停行始 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Row  'Rangeの引数はお好みで
    年停行終 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows(Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows.Count).Row 'Rangeの引数はお好みで
    作業内容列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("作業内容", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    実施箇所列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("実施箇所", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    開始日時列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求期間　開始日時", LookAt:=xlWhole).Column   'Rangeの引数は必要に応じて変更
    要求時注釈列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求時注釈", LookAt:=xlWhole).Column
    停止設備列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("停止設備および線路名　停止設備１", LookAt:=xlWhole).Column
    
    For i = 行始 To 行終
    
        If Rows(i).Hidden = False And Cells(i, 電気所等名称列).Value Like "*" & "　" & "*" Then  '可視セルかつ電気所名に全角スペースが含まれる場合のみ処理を行う

            電気所等名称 = Left(Cells(i, 電気所等名称列).Value, InStr(Cells(i, 電気所等名称列).Value, "　") - 1)    '左端から電気所名のみを切り出す
            If Right(電気所等名称, 1) = "開" Then
                電気所等名称 = 電気所等名称 & "閉所"
            Else
                電気所等名称 = 電気所等名称 & "電所"
            End If
            
            保全計画テキスト = StrConv(Cells(i, 保全計画テキスト列).Value, vbWide)    '全角にする
            保全計画テキスト = StrConv(保全計画テキスト, vbUpperCase) '大文字にする
            保全計画テキスト = Replace(保全計画テキスト, "　ＬＰＤ", "ＬＰＤ")
    
            If 保全計画テキスト Like "*" & 機器種類 & "*" And (Cells(i, クラス列).Value = "X_APT_") Then
    
                切始 = InStr(保全計画テキスト, 機器種類)
                
                For j = 切始 - 1 To 1 Step -1
                    If Mid(保全計画テキスト, j, 1) = "　" Then
                        切出し文字 = Mid(保全計画テキスト, j + 1, 切始 - 1 - j)
                        Exit For
                        
                    End If
                Next j

                For k = 年停行始 To 年停行終
                
                    If Workbooks(検索先ワークブック名).Sheets(1).Rows(k).Hidden = False Then  '可視セルのみ処理を行う
                
                        作業内容 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Value, vbWide)  '全角にする
                        作業内容 = StrConv(作業内容, vbUpperCase) '大文字にする
                        
                        要求時注釈 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Value, vbWide)  '全角にする
                        要求時注釈 = StrConv(要求時注釈, vbUpperCase) '大文字にする
                        
                        実施個所 = Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 実施箇所列).Value
                        
                        停止設備 = Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 停止設備列).Value
                        停止設備 = StrConv(停止設備, vbWide)  '全角にする
                        停止設備 = StrConv(停止設備, vbUpperCase) '大文字にする
                
                        If 実施個所 = 電気所等名称 And 作業内容 Like "*" & 機器種類 & "*" And 停止設備 Like "*" & 切出し文字 & "*" Then    '実施個所と作業内容が一致していれば着色
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                            
                        ElseIf 実施個所 = 電気所等名称 And 要求時注釈 Like "*" & 機器種類 & "*" And 停止設備 Like "*" & 切出し文字 & "*" Then
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                        
                        End If
                    
                    End If
                    
                    DoEvents
                
                Next k

            End If

        End If

        Application.StatusBar = "処理実行中．．．" & Round(i / 行終 * 100, 0) & "%"
        
    Next i
    
    Application.StatusBar = False
    
End Sub
Sub 保全計画と設備停止計画のチェックCB()
    '検索元と検索先のブックを開いておく
    '検索元のブックをアクティブにし、カレントリージョンが効く位置でマクロを実行する
    'Workbooksからのオブジェクト指定が無いものは保全計画側の処理です
    'クラスおよび設備停止有無でフィルタをかけてから実行するとより確実です（可視セルのみ処理対象のため）
    Dim tmp As Variant
    Dim 検索先ワークブック名 As String
    Dim 検索元ワークブック名 As String
    Dim 電気所等名称列 As Integer
    Dim 保全計画テキスト列 As Integer
    Dim クラス列 As Integer
    Dim 月範囲 As Range
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 列始 As Integer
    Dim 列終 As Integer
    Dim 年停行始 As Long
    Dim 年停行終 As Long
    Dim 作業内容列 As Integer
    Dim 実施箇所列 As Integer
    Dim 開始日時列 As Integer
    Dim i As Long
    Dim 電気所等名称 As String
    Dim 保全計画テキスト As String
    Dim 切始 As Integer
    Dim 切出し文字 As String
    Dim 作業内容 As String
    Dim 実施個所 As String
    Dim 月 As Integer
    Dim 要求時注釈列 As Integer
    Dim 要求時注釈 As String
    
    機器種類 = "ＣＢ"

    検索先ワークブック名 = InputBox("設備停止側のブック名を拡張子を含めて入力して下さい。")
    検索元ワークブック名 = ActiveWorkbook.Name
    電気所等名称列 = 2  '必要に応じて変更
    保全計画テキスト列 = 6  '必要に応じて変更
    クラス列 = 32   '必要に応じて変更
    Set 月範囲 = Range("M3:X3") '必要に応じて変更

    行始 = ActiveCell.CurrentRegion.Row
    行終 = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row
    列始 = ActiveCell.CurrentRegion.Column
    列終 = ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column
    年停行始 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Row  'Rangeの引数はお好みで
    年停行終 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows(Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows.Count).Row 'Rangeの引数はお好みで
    作業内容列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("作業内容", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    実施箇所列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("実施箇所", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    開始日時列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求期間　開始日時", LookAt:=xlWhole).Column   'Rangeの引数は必要に応じて変更
    要求時注釈列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求時注釈", LookAt:=xlWhole).Column
    
    For i = 行始 To 行終
    
        If Rows(i).Hidden = False And Cells(i, 電気所等名称列).Value Like "*" & "　" & "*" Then  '可視セルかつ電気所名に全角スペースが含まれる場合のみ処理を行う

            電気所等名称 = Left(Cells(i, 電気所等名称列).Value, InStr(Cells(i, 電気所等名称列).Value, "　") - 1)    '左端から電気所名のみを切り出す
            If Right(電気所等名称, 1) = "開" Then
                電気所等名称 = 電気所等名称 & "閉所"
            Else
                電気所等名称 = 電気所等名称 & "電所"
            End If
            
            保全計画テキスト = StrConv(Cells(i, 保全計画テキスト列).Value, vbWide)    '全角にする
            保全計画テキスト = StrConv(保全計画テキスト, vbUpperCase) '大文字にする
            保全計画テキスト = Replace(保全計画テキスト, "ＣＢ　", "ＣＢ")
    
            If 保全計画テキスト Like "*" & 機器種類 & "*" Then
    
                切始 = InStr(保全計画テキスト, 機器種類)
                
                    For j = 切始 To Len(保全計画テキスト)
                        If Mid(保全計画テキスト, j, 1) Like "[!０-９Ａ-Ｚ－]" Then
                            切出し文字 = Mid(保全計画テキスト, 切始, j - 切始)
                            Exit For
            
                        End If
                    Next j

                For k = 年停行始 To 年停行終
                
                    If Workbooks(検索先ワークブック名).Sheets(1).Rows(k).Hidden = False Then  '可視セルのみ処理を行う
                
                        作業内容 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Value, vbWide)  '全角にする
                        作業内容 = StrConv(作業内容, vbUpperCase) '大文字にする
                        
                        要求時注釈 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Value, vbWide)  '全角にする
                        要求時注釈 = StrConv(要求時注釈, vbUpperCase) '大文字にする
                        
                        実施個所 = Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 実施箇所列).Value
                
                        If 実施個所 = 電気所等名称 And 作業内容 Like "*" & 切出し文字 & "*" Then    '実施個所と作業内容が一致していれば着色
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                            
                        ElseIf 実施個所 = 電気所等名称 And 要求時注釈 Like "*" & 切出し文字 & "*" Then
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                        
                        End If
                    
                    End If
                    
                    DoEvents
                
                Next k

            End If

        End If

        Application.StatusBar = "処理実行中．．．" & Round(i / 行終 * 100, 0) & "%"
        
    Next i
    
    Application.StatusBar = False
    
End Sub
Sub 保全計画と設備停止計画のチェックLS()
    '検索元と検索先のブックを開いておく
    '検索元のブックをアクティブにし、カレントリージョンが効く位置でマクロを実行する
    'Workbooksからのオブジェクト指定が無いものは保全計画側の処理です
    'クラスおよび設備停止有無でフィルタをかけてから実行するとより確実です（可視セルのみ処理対象のため）
    Dim tmp As Variant
    Dim 検索先ワークブック名 As String
    Dim 検索元ワークブック名 As String
    Dim 電気所等名称列 As Integer
    Dim 保全計画テキスト列 As Integer
    Dim クラス列 As Integer
    Dim 月範囲 As Range
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 列始 As Integer
    Dim 列終 As Integer
    Dim 年停行始 As Long
    Dim 年停行終 As Long
    Dim 作業内容列 As Integer
    Dim 実施箇所列 As Integer
    Dim 開始日時列 As Integer
    Dim i As Long
    Dim 電気所等名称 As String
    Dim 保全計画テキスト As String
    Dim 切始 As Integer
    Dim 切出し文字 As String
    Dim 作業内容 As String
    Dim 実施個所 As String
    Dim 月 As Integer
    Dim 要求時注釈列 As Integer
    Dim 要求時注釈 As String
    
    機器種類 = "ＬＳ"

    検索先ワークブック名 = InputBox("設備停止側のブック名を拡張子を含めて入力して下さい。")
    検索元ワークブック名 = ActiveWorkbook.Name
    電気所等名称列 = 2  '必要に応じて変更
    保全計画テキスト列 = 6  '必要に応じて変更
    クラス列 = 32   '必要に応じて変更
    Set 月範囲 = Range("M3:X3") '必要に応じて変更

    行始 = ActiveCell.CurrentRegion.Row
    行終 = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row
    列始 = ActiveCell.CurrentRegion.Column
    列終 = ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column
    年停行始 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Row  'Rangeの引数はお好みで
    年停行終 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows(Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows.Count).Row 'Rangeの引数はお好みで
    作業内容列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("作業内容", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    実施箇所列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("実施箇所", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    開始日時列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求期間　開始日時", LookAt:=xlWhole).Column   'Rangeの引数は必要に応じて変更
    要求時注釈列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求時注釈", LookAt:=xlWhole).Column
    
    For i = 行始 To 行終
    
        If Rows(i).Hidden = False And Cells(i, 電気所等名称列).Value Like "*" & "　" & "*" Then  '可視セルかつ電気所名に全角スペースが含まれる場合のみ処理を行う

            電気所等名称 = Left(Cells(i, 電気所等名称列).Value, InStr(Cells(i, 電気所等名称列).Value, "　") - 1)    '左端から電気所名のみを切り出す
            If Right(電気所等名称, 1) = "開" Then
                電気所等名称 = 電気所等名称 & "閉所"
            Else
                電気所等名称 = 電気所等名称 & "電所"
            End If
            
            保全計画テキスト = StrConv(Cells(i, 保全計画テキスト列).Value, vbWide)    '全角にする
            保全計画テキスト = StrConv(保全計画テキスト, vbUpperCase) '大文字にする
            保全計画テキスト = Replace(保全計画テキスト, "ＬＳ　", "ＬＳ")
    
            If 保全計画テキスト Like "*" & 機器種類 & "*" Then
    
                切始 = InStr(保全計画テキスト, 機器種類)
                
                    For j = 切始 To Len(保全計画テキスト)
                        If Mid(保全計画テキスト, j, 1) Like "[!０-９Ａ-Ｚ－]" Then
                            切出し文字 = Mid(保全計画テキスト, 切始, j - 切始)
                            Exit For
            
                        End If
                    Next j

                For k = 年停行始 To 年停行終
                
                    If Workbooks(検索先ワークブック名).Sheets(1).Rows(k).Hidden = False Then  '可視セルのみ処理を行う
                
                        作業内容 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Value, vbWide)  '全角にする
                        作業内容 = StrConv(作業内容, vbUpperCase) '大文字にする
                        
                        要求時注釈 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Value, vbWide)  '全角にする
                        要求時注釈 = StrConv(要求時注釈, vbUpperCase) '大文字にする
                        
                        実施個所 = Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 実施箇所列).Value
                
                        If 実施個所 = 電気所等名称 And 作業内容 Like "*" & 切出し文字 & "*" Then    '実施個所と作業内容が一致していれば着色
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                            
                        ElseIf 実施個所 = 電気所等名称 And 要求時注釈 Like "*" & 切出し文字 & "*" Then
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                        
                        End If
                    
                    End If
                    
                    DoEvents
                
                Next k

            End If

        End If

        Application.StatusBar = "処理実行中．．．" & Round(i / 行終 * 100, 0) & "%"
        
    Next i
    
    Application.StatusBar = False
    
End Sub
Sub 保全計画と設備停止計画のチェックTR()
    '検索元と検索先のブックを開いておく
    '検索元のブックをアクティブにし、カレントリージョンが効く位置でマクロを実行する
    'Workbooksからのオブジェクト指定が無いものは保全計画側の処理です
    'クラスおよび設備停止有無でフィルタをかけてから実行するとより確実です（可視セルのみ処理対象のため）
    Dim tmp As Variant
    Dim 検索先ワークブック名 As String
    Dim 検索元ワークブック名 As String
    Dim 電気所等名称列 As Integer
    Dim 保全計画テキスト列 As Integer
    Dim クラス列 As Integer
    Dim 月範囲 As Range
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 列始 As Integer
    Dim 列終 As Integer
    Dim 年停行始 As Long
    Dim 年停行終 As Long
    Dim 作業内容列 As Integer
    Dim 実施箇所列 As Integer
    Dim 開始日時列 As Integer
    Dim i As Long
    Dim 電気所等名称 As String
    Dim 保全計画テキスト As String
    Dim 切始 As Integer
    Dim 切出し文字 As String
    Dim 作業内容 As String
    Dim 実施個所 As String
    Dim 月 As Integer
    
    機器種類 = "ＴＲ"

    検索先ワークブック名 = InputBox("設備停止側のブック名を拡張子を含めて入力して下さい。")
    検索元ワークブック名 = ActiveWorkbook.Name
    電気所等名称列 = 2  '必要に応じて変更
    保全計画テキスト列 = 6  '必要に応じて変更
    クラス列 = 32   '必要に応じて変更
    Set 月範囲 = Range("M3:X3") '必要に応じて変更

    行始 = ActiveCell.CurrentRegion.Row
    行終 = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row
    列始 = ActiveCell.CurrentRegion.Column
    列終 = ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column
    年停行始 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Row  'Rangeの引数はお好みで
    年停行終 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows(Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows.Count).Row 'Rangeの引数はお好みで
    作業内容列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("作業内容", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    実施箇所列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("実施箇所", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    開始日時列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求期間　開始日時", LookAt:=xlWhole).Column   'Rangeの引数は必要に応じて変更
    
    For i = 行始 To 行終
    
        If Rows(i).Hidden = False And Cells(i, 電気所等名称列).Value Like "*" & "　" & "*" Then  '可視セルかつ電気所名に全角スペースが含まれる場合のみ処理を行う

            電気所等名称 = Left(Cells(i, 電気所等名称列).Value, InStr(Cells(i, 電気所等名称列).Value, "　") - 1)    '左端から電気所名のみを切り出す
            If Right(電気所等名称, 1) = "開" Then
                電気所等名称 = 電気所等名称 & "閉所"
            Else
                電気所等名称 = 電気所等名称 & "電所"
            End If
            
            保全計画テキスト = StrConv(Cells(i, 保全計画テキスト列).Value, vbWide)    '全角にする
            保全計画テキスト = StrConv(保全計画テキスト, vbUpperCase) '大文字にする
    
            If 保全計画テキスト Like "*" & 機器種類 & "*" Then
    
                切始 = InStr(保全計画テキスト, 機器種類)
                
                If 機器種類 = "ＣＢ" Or 機器種類 = "ＬＳ" Then  'ＣＢ、ＬＳの場合の切り出し処理
                    For j = 切始 To Len(保全計画テキスト)
                        If Mid(保全計画テキスト, j, 1) Like "[!０-９Ａ-Ｚ－]" Then
                            切出し文字 = Mid(保全計画テキスト, 切始, j - 切始)
                            Exit For
            
                        End If
                    Next j
                ElseIf 機器種類 = "ＴＲ" And (Cells(i, クラス列).Value = "X_ATR_" Or Cells(i, クラス列).Value = "X_AGTR" Or Cells(i, クラス列).Value = "X_ASTR") Then 'ＴＲの場合の切り出し処理
                    For j = 切始 To 1 Step -1
                        If Mid(保全計画テキスト, j, 1) Like "[!０-９Ａ-Ｚ－]" Then
                            切出し文字 = Mid(保全計画テキスト, j + 1, 切始 + Len(機器種類) - 1 - j)
                            Exit For
            
                        End If
                    Next j
                Else
                    GoTo 何もしない
                    
                End If

                For k = 年停行始 To 年停行終
                
                    If Workbooks(検索先ワークブック名).Sheets(1).Rows(k).Hidden = False Then  '可視セルのみ処理を行う
                
                        作業内容 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Value, vbWide)  '全角にする
                        作業内容 = StrConv(作業内容, vbUpperCase) '大文字にする
                        実施個所 = Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 実施箇所列).Value
                
                        If 実施個所 = 電気所等名称 And 作業内容 Like "*" & 切出し文字 & "*" Then    '実施個所と作業内容が一致していれば着色
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                        
                        End If
                    
                    End If
                    
                    DoEvents
                
                Next k

            End If

        End If
        
何もしない:

        Application.StatusBar = "処理実行中．．．" & Round(i / 行終 * 100, 0) & "%"
        
    Next i
    
    Application.StatusBar = False
    
End Sub
Sub 保全計画と設備停止計画のチェックTR要求時注釈側()
    '検索元と検索先のブックを開いておく
    '検索元のブックをアクティブにし、カレントリージョンが効く位置でマクロを実行する
    'Workbooksからのオブジェクト指定が無いものは保全計画側の処理です
    'クラスおよび設備停止有無でフィルタをかけてから実行するとより確実です（可視セルのみ処理対象のため）
    '要求時注釈欄の、作業内容～・の間を切り取ってチェックします
    '・が意図せぬ所にあると不完全なチェックとなるため、正規版の補助として使用する
    '正規版で着色できなかったものにフィルタをかけてから実行すると効率的です
    'LSやCBと数字の間に半角スペースがあるものは一括置換で削除してから実行する
    '処理対象の列が非表示になっているとエラーになる
    Dim tmp As Variant
    Dim 検索先ワークブック名 As String
    Dim 検索元ワークブック名 As String
    Dim 電気所等名称列 As Integer
    Dim 保全計画テキスト列 As Integer
    Dim クラス列 As Integer
    Dim 月範囲 As Range
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 列始 As Integer
    Dim 列終 As Integer
    Dim 年停行始 As Long
    Dim 年停行終 As Long
    Dim 作業内容列 As Integer   '要求時注釈用として流用
    Dim 実施箇所列 As Integer
    Dim 開始日時列 As Integer
    Dim i As Long
    Dim 電気所等名称 As String
    Dim 保全計画テキスト As String
    Dim 切始 As Integer
    Dim 切出し文字 As String
    Dim 作業内容 As String  '要求時注釈用として流用
    Dim 実施個所 As String
    Dim 月 As Integer
    Dim 要求切始 As Integer
    Dim 要求切終 As Integer
    
    機器種類 = "ＴＲ"

    検索先ワークブック名 = InputBox("設備停止側のブック名を拡張子を含めて入力して下さい。")
    検索元ワークブック名 = ActiveWorkbook.Name
    電気所等名称列 = 2  '必要に応じて変更
    保全計画テキスト列 = 6  '必要に応じて変更
    クラス列 = 32   '必要に応じて変更
    Set 月範囲 = Range("M3:X3") '必要に応じて変更

    行始 = ActiveCell.CurrentRegion.Row
    行終 = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row
    列始 = ActiveCell.CurrentRegion.Column
    列終 = ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column
    年停行始 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Row  'Rangeの引数はお好みで
    年停行終 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows(Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows.Count).Row 'Rangeの引数はお好みで
    作業内容列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求時注釈", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更・要求時注釈用として流用
    実施箇所列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("実施箇所", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    開始日時列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求期間　開始日時", LookAt:=xlWhole).Column   'Rangeの引数は必要に応じて変更
    
    For i = 行始 To 行終
    
        If Rows(i).Hidden = False And Cells(i, 電気所等名称列).Value Like "*" & "　" & "*" And Cells(i, 保全計画テキスト列).Interior.Color = 16777215 Then  '可視セルかつ電気所名に全角スペースが含まれる場合のみ処理を行う

            電気所等名称 = Left(Cells(i, 電気所等名称列).Value, InStr(Cells(i, 電気所等名称列).Value, "　") - 1)    '左端から電気所名のみを切り出す
            If Right(電気所等名称, 1) = "開" Then
                電気所等名称 = 電気所等名称 & "閉所"
            Else
                電気所等名称 = 電気所等名称 & "電所"
            End If
            
            保全計画テキスト = StrConv(Cells(i, 保全計画テキスト列).Value, vbWide)    '全角にする
            保全計画テキスト = StrConv(保全計画テキスト, vbUpperCase) '大文字にする
    
            If 保全計画テキスト Like "*" & 機器種類 & "*" Then
    
                切始 = InStr(保全計画テキスト, 機器種類)
                
                If 機器種類 = "ＣＢ" Or 機器種類 = "ＬＳ" Then  'ＣＢ、ＬＳの場合の切り出し処理
                    For j = 切始 To Len(保全計画テキスト)
                        If Mid(保全計画テキスト, j, 1) Like "[!０-９Ａ-Ｚ－]" Then
                            切出し文字 = Mid(保全計画テキスト, 切始, j - 切始)
                            Exit For
            
                        End If
                    Next j
                ElseIf 機器種類 = "ＴＲ" And (Cells(i, クラス列).Value = "X_ATR_" Or Cells(i, クラス列).Value = "X_AGTR" Or Cells(i, クラス列).Value = "X_ASTR") Then 'ＴＲの場合の切り出し処理
                    For j = 切始 To 1 Step -1
                        If Mid(保全計画テキスト, j, 1) Like "[!０-９Ａ-Ｚ－]" Then
                            切出し文字 = Mid(保全計画テキスト, j + 1, 切始 + Len(機器種類) - 1 - j)
                            Exit For
            
                        End If
                    Next j
                Else
                    GoTo 何もしない
                    
                End If

                For k = 年停行始 To 年停行終
                
                    If Workbooks(検索先ワークブック名).Sheets(1).Rows(k).Hidden = False Then  '可視セルのみ処理を行う
                    
                        作業内容 = Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Value '要求時注釈用として追記
                        要求切始 = InStr(作業内容, "作業内容")  '要求時注釈用として追記
                        If 要求切始 <> 0 Then   '要求時注釈用として追記
                            要求切終 = InStr(要求切始, 作業内容, "・")  '要求時注釈用として追記
                            If 要求切終 <> 0 Then   '要求時注釈用として追記
                                作業内容 = Mid(作業内容, 要求切始, 要求切終 - 要求切始) '要求時注釈用として追記
                            Else
                                作業内容 = Right(作業内容, Len(作業内容) - 要求切始)    '要求時注釈用として追記
                            End If

                            作業内容 = StrConv(作業内容, vbWide)  '全角にする・要求時注釈用として引数を変更
                            作業内容 = StrConv(作業内容, vbUpperCase) '大文字にする
                            実施個所 = Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 実施箇所列).Value
                
                            If 実施個所 = 電気所等名称 And 作業内容 Like "*" & 切出し文字 & "*" Then    '実施個所と作業内容が一致していれば着色
                                月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                                Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Interior.Color = 65535
                                Cells(i, 保全計画テキスト列).Interior.Color = 65535
                                Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                                Exit For
                        
                            End If
                            
                        End If
                    
                    End If
                    
                    DoEvents
                
                Next k

            End If

        End If
        
何もしない:

        Application.StatusBar = "処理実行中．．．" & Round(i / 行終 * 100, 0) & "%"
        
    Next i
    
    Application.StatusBar = False
    
End Sub
Sub 保全計画と設備停止計画のチェックTR盤()
    '検索元と検索先のブックを開いておく
    '検索元のブックをアクティブにし、カレントリージョンが効く位置でマクロを実行する
    'Workbooksからのオブジェクト指定が無いものは保全計画側の処理です
    'クラスおよび設備停止有無でフィルタをかけてから実行するとより確実です（可視セルのみ処理対象のため）
    Dim tmp As Variant
    Dim 検索先ワークブック名 As String
    Dim 検索元ワークブック名 As String
    Dim 電気所等名称列 As Integer
    Dim 保全計画テキスト列 As Integer
    Dim クラス列 As Integer
    Dim 月範囲 As Range
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 列始 As Integer
    Dim 列終 As Integer
    Dim 年停行始 As Long
    Dim 年停行終 As Long
    Dim 作業内容列 As Integer
    Dim 実施箇所列 As Integer
    Dim 開始日時列 As Integer
    Dim i As Long
    Dim 電気所等名称 As String
    Dim 保全計画テキスト As String
    Dim 切始 As Integer
    Dim 切出し文字 As String
    Dim 作業内容 As String
    Dim 実施個所 As String
    Dim 月 As Integer
    Dim 要求時注釈列 As Integer
    Dim 要求時注釈 As String
    Dim 設備名称列 As Integer
    
    機器種類 = "変圧器盤"

    検索先ワークブック名 = InputBox("設備停止側のブック名を拡張子を含めて入力して下さい。")
    検索元ワークブック名 = ActiveWorkbook.Name
    電気所等名称列 = 2  '必要に応じて変更
    設備名称列 = 4
    保全計画テキスト列 = 6  '必要に応じて変更
    クラス列 = 32   '必要に応じて変更
    Set 月範囲 = Range("M3:X3") '必要に応じて変更

    行始 = ActiveCell.CurrentRegion.Row
    行終 = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row
    列始 = ActiveCell.CurrentRegion.Column
    列終 = ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column
    年停行始 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Row  'Rangeの引数はお好みで
    年停行終 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows(Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows.Count).Row 'Rangeの引数はお好みで
    作業内容列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("作業内容", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    実施箇所列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("実施箇所", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    開始日時列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求期間　開始日時", LookAt:=xlWhole).Column   'Rangeの引数は必要に応じて変更
    要求時注釈列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求時注釈", LookAt:=xlWhole).Column
    
    For i = 行始 To 行終
    
        If Rows(i).Hidden = False And Cells(i, 電気所等名称列).Value Like "*" & "　" & "*" Then  '可視セルかつ電気所名に全角スペースが含まれる場合のみ処理を行う

            電気所等名称 = Left(Cells(i, 電気所等名称列).Value, InStr(Cells(i, 電気所等名称列).Value, "　") - 1)    '左端から電気所名のみを切り出す
            If Right(電気所等名称, 1) = "開" Then
                電気所等名称 = 電気所等名称 & "閉所"
            Else
                電気所等名称 = 電気所等名称 & "電所"
            End If
            
            保全計画テキスト = StrConv(Cells(i, 設備名称列).Value, vbWide)    '全角にする
            保全計画テキスト = StrConv(保全計画テキスト, vbUpperCase) '大文字にする
    
            If 保全計画テキスト Like "*" & 機器種類 & "*" And (Cells(i, クラス列).Value = "X_ASB_") Then
    
                切出し文字 = Replace(保全計画テキスト, "号", "ＴＲ")
                切出し文字 = Replace(切出し文字, "配変", "")
                切出し文字 = Replace(切出し文字, "配", "")
                切出し文字 = Replace(切出し文字, "連変", "")
                切出し文字 = Replace(切出し文字, "変圧器", "")

                For k = 年停行始 To 年停行終
                
                    If Workbooks(検索先ワークブック名).Sheets(1).Rows(k).Hidden = False Then  '可視セルのみ処理を行う
                
                        作業内容 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Value, vbWide)  '全角にする
                        作業内容 = StrConv(作業内容, vbUpperCase) '大文字にする
                        作業内容 = Replace(作業内容, "号変圧器盤", "ＴＲ盤")
                        
                        要求時注釈 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Value, vbWide)  '全角にする
                        要求時注釈 = StrConv(要求時注釈, vbUpperCase) '大文字にする
                        要求時注釈 = Replace(要求時注釈, "号変圧器盤", "ＴＲ盤")
                        
                        実施個所 = Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 実施箇所列).Value
                
                        If 実施個所 = 電気所等名称 And 作業内容 Like "*" & 切出し文字 & "*" Then    '実施個所と作業内容が一致していれば着色
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                            
                        ElseIf 実施個所 = 電気所等名称 And 要求時注釈 Like "*" & 切出し文字 & "*" Then
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                        
                        End If
                    
                    End If
                    
                    DoEvents
                
                Next k

            End If

        End If

        Application.StatusBar = "処理実行中．．．" & Round(i / 行終 * 100, 0) & "%"
        
    Next i
    
    Application.StatusBar = False
    
End Sub
Sub 保全計画と設備停止計画のチェックPWC()
    '検索元と検索先のブックを開いておく
    '検索元のブックをアクティブにし、カレントリージョンが効く位置でマクロを実行する
    'Workbooksからのオブジェクト指定が無いものは保全計画側の処理です
    'クラスおよび設備停止有無でフィルタをかけてから実行するとより確実です（可視セルのみ処理対象のため）
    Dim tmp As Variant
    Dim 検索先ワークブック名 As String
    Dim 検索元ワークブック名 As String
    Dim 電気所等名称列 As Integer
    Dim 保全計画テキスト列 As Integer
    Dim クラス列 As Integer
    Dim 月範囲 As Range
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 列始 As Integer
    Dim 列終 As Integer
    Dim 年停行始 As Long
    Dim 年停行終 As Long
    Dim 作業内容列 As Integer
    Dim 実施箇所列 As Integer
    Dim 開始日時列 As Integer
    Dim i As Long
    Dim 電気所等名称 As String
    Dim 保全計画テキスト As String
    Dim 切始 As Integer
    Dim 切出し文字 As String
    Dim 作業内容 As String
    Dim 実施個所 As String
    Dim 月 As Integer
    Dim 要求時注釈列 As Integer
    Dim 要求時注釈 As String
    Dim 設備名称列 As Integer
    
    機器種類 = "電力ケーブル"

    検索先ワークブック名 = InputBox("設備停止側のブック名を拡張子を含めて入力して下さい。")
    検索元ワークブック名 = ActiveWorkbook.Name
    電気所等名称列 = 2  '必要に応じて変更
    設備名称列 = 4
    保全計画テキスト列 = 6  '必要に応じて変更
    クラス列 = 32   '必要に応じて変更
    Set 月範囲 = Range("M3:X3") '必要に応じて変更

    行始 = ActiveCell.CurrentRegion.Row
    行終 = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row
    列始 = ActiveCell.CurrentRegion.Column
    列終 = ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column
    年停行始 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Row  'Rangeの引数はお好みで
    年停行終 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows(Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows.Count).Row 'Rangeの引数はお好みで
    作業内容列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("作業内容", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    実施箇所列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("実施箇所", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    開始日時列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求期間　開始日時", LookAt:=xlWhole).Column   'Rangeの引数は必要に応じて変更
    要求時注釈列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求時注釈", LookAt:=xlWhole).Column
    
    For i = 行始 To 行終
    
        If Rows(i).Hidden = False And Cells(i, 電気所等名称列).Value Like "*" & "　" & "*" Then  '可視セルかつ電気所名に全角スペースが含まれる場合のみ処理を行う

            電気所等名称 = Left(Cells(i, 電気所等名称列).Value, InStr(Cells(i, 電気所等名称列).Value, "　") - 1)    '左端から電気所名のみを切り出す
            If Right(電気所等名称, 1) = "開" Then
                電気所等名称 = 電気所等名称 & "閉所"
            Else
                電気所等名称 = 電気所等名称 & "電所"
            End If
            
            保全計画テキスト = StrConv(Cells(i, 設備名称列).Value, vbWide)    '全角にする
            保全計画テキスト = StrConv(保全計画テキスト, vbUpperCase) '大文字にする
    
            If 保全計画テキスト Like "*" & 機器種類 & "*" And (Cells(i, クラス列).Value = "X_APWC") Then
    
                切出し文字 = Replace(保全計画テキスト, "配変", "")
                切出し文字 = Replace(切出し文字, "配", "")

                For k = 年停行始 To 年停行終
                
                    If Workbooks(検索先ワークブック名).Sheets(1).Rows(k).Hidden = False Then  '可視セルのみ処理を行う
                
                        作業内容 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Value, vbWide)  '全角にする
                        作業内容 = StrConv(作業内容, vbUpperCase) '大文字にする
                        作業内容 = Replace(作業内容, "ＰＷＣ", 機器種類)
                        作業内容 = Replace(作業内容, "一次", "１次")
                        作業内容 = Replace(作業内容, "二次", "２次")
                        
                        要求時注釈 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Value, vbWide)  '全角にする
                        要求時注釈 = StrConv(要求時注釈, vbUpperCase) '大文字にする
                        要求時注釈 = Replace(要求時注釈, "ＰＷＣ", 機器種類)
                        要求時注釈 = Replace(要求時注釈, "一次", "１次")
                        要求時注釈 = Replace(要求時注釈, "二次", "２次")
                        
                        実施個所 = Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 実施箇所列).Value
                
                        If 実施個所 = 電気所等名称 And 作業内容 Like "*" & 切出し文字 & "*" Then    '実施個所と作業内容が一致していれば着色
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                            
                        ElseIf 実施個所 = 電気所等名称 And 要求時注釈 Like "*" & 切出し文字 & "*" Then
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                        
                        End If
                    
                    End If
                    
                    DoEvents
                
                Next k

            End If

        End If

        Application.StatusBar = "処理実行中．．．" & Round(i / 行終 * 100, 0) & "%"
        
    Next i
    
    Application.StatusBar = False
    
End Sub
Sub 保全計画と設備停止計画のチェックMOF()
    '検索元と検索先のブックを開いておく
    '検索元のブックをアクティブにし、カレントリージョンが効く位置でマクロを実行する
    'Workbooksからのオブジェクト指定が無いものは保全計画側の処理です
    'クラスおよび設備停止有無でフィルタをかけてから実行するとより確実です（可視セルのみ処理対象のため）
    Dim tmp As Variant
    Dim 検索先ワークブック名 As String
    Dim 検索元ワークブック名 As String
    Dim 電気所等名称列 As Integer
    Dim 保全計画テキスト列 As Integer
    Dim クラス列 As Integer
    Dim 月範囲 As Range
    Dim 行始 As Long
    Dim 行終 As Long
    Dim 列始 As Integer
    Dim 列終 As Integer
    Dim 年停行始 As Long
    Dim 年停行終 As Long
    Dim 作業内容列 As Integer
    Dim 実施箇所列 As Integer
    Dim 開始日時列 As Integer
    Dim i As Long
    Dim 電気所等名称 As String
    Dim 保全計画テキスト As String
    Dim 切始 As Integer
    Dim 切出し文字 As String
    Dim 作業内容 As String
    Dim 実施個所 As String
    Dim 月 As Integer
    Dim 要求時注釈列 As Integer
    Dim 要求時注釈 As String
    Dim 設備名称列 As Integer

    検索先ワークブック名 = InputBox("設備停止側のブック名を拡張子を含めて入力して下さい。")
    検索元ワークブック名 = ActiveWorkbook.Name
    電気所等名称列 = 2  '必要に応じて変更
    設備名称列 = 4
    保全計画テキスト列 = 6  '必要に応じて変更
    クラス列 = 32   '必要に応じて変更
    Set 月範囲 = Range("M3:X3") '必要に応じて変更

    行始 = ActiveCell.CurrentRegion.Row
    行終 = ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row
    列始 = ActiveCell.CurrentRegion.Column
    列終 = ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column
    年停行始 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Row  'Rangeの引数はお好みで
    年停行終 = Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows(Workbooks(検索先ワークブック名).Sheets(1).Range("E5").CurrentRegion.Rows.Count).Row 'Rangeの引数はお好みで
    作業内容列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("作業内容", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    実施箇所列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("実施箇所", LookAt:=xlWhole).Column 'Rangeの引数は必要に応じて変更
    開始日時列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求期間　開始日時", LookAt:=xlWhole).Column   'Rangeの引数は必要に応じて変更
    要求時注釈列 = Workbooks(検索先ワークブック名).Sheets(1).Range("A1:FK1").Find("要求時注釈", LookAt:=xlWhole).Column
    
    For i = 行始 To 行終
    
        If Rows(i).Hidden = False And Cells(i, 電気所等名称列).Value Like "*" & "　" & "*" Then  '可視セルかつ電気所名に全角スペースが含まれる場合のみ処理を行う

            電気所等名称 = Left(Cells(i, 電気所等名称列).Value, InStr(Cells(i, 電気所等名称列).Value, "　") - 1)    '左端から電気所名のみを切り出す
            If Right(電気所等名称, 1) = "開" Then
                電気所等名称 = 電気所等名称 & "閉所"
            Else
                電気所等名称 = 電気所等名称 & "電所"
            End If
    
            If Cells(i, クラス列).Value = "X_AMPC" Then

                For k = 年停行始 To 年停行終
                
                    If Workbooks(検索先ワークブック名).Sheets(1).Rows(k).Hidden = False Then  '可視セルのみ処理を行う
                
                        作業内容 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Value, vbWide)  '全角にする
                        作業内容 = StrConv(作業内容, vbUpperCase) '大文字にする
                        作業内容 = Replace(作業内容, "計量装置", "ＭＯＦ")
                        
                        要求時注釈 = StrConv(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Value, vbWide)  '全角にする
                        要求時注釈 = StrConv(要求時注釈, vbUpperCase) '大文字にする
                        要求時注釈 = Replace(要求時注釈, "計量装置", "ＭＯＦ")
                        
                        実施個所 = Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 実施箇所列).Value
                        
                        保全計画テキスト = StrConv(Cells(i, 保全計画テキスト列).Value, vbWide)    '全角にする
                
                        If 保全計画テキスト Like "*" & 実施個所 & "*" And 作業内容 Like "*ＭＯＦ*" Then    '実施個所と作業内容が一致していれば着色
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 作業内容列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                            
                        ElseIf 保全計画テキスト Like "*" & 実施個所 & "*" And 要求時注釈 Like "*ＭＯＦ*" Then
                            月 = Month(Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 開始日時列).Value)    '設備停止の計画月を抽出
                            
                            Workbooks(検索先ワークブック名).Sheets(1).Cells(k, 要求時注釈列).Interior.Color = 65535
                            Cells(i, 保全計画テキスト列).Interior.Color = 65535
                            Cells(i, 月範囲.Find(CStr(月), LookAt:=xlWhole).Column).Interior.Color = 65535  '保全計画に設備停止の計画月を着色

                            Exit For
                        
                        End If
                    
                    End If
                    
                    DoEvents
                
                Next k

            End If

        End If

        Application.StatusBar = "処理実行中．．．" & Round(i / 行終 * 100, 0) & "%"
        
    Next i
    
    Application.StatusBar = False
    
End Sub