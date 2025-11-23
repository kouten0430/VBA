Sub スケジュールへインポート用のデータを作成()
    Dim i As Long
    Dim 日付行 As Long
    Dim 名前列 As Long
    Dim 月初 As Long
    Dim 月末 As Long
    Dim 範囲 As Range
    Dim tmp As Worksheet
    Dim 転記先シート As Worksheet
    Dim myRange As Range
    Dim 同行 As String
    Dim j As Long
    Dim 一致 As Range
    Dim 予定 As Variant
    
    '---ここから行番号・列番号を取得する処理---
    
    For i = 1 To Cells.Find("組織スケジュール", LookAt:=xlWhole).Row
        If TypeName(Cells(i, "S").Value) = "Date" Then
            日付行 = i
            Exit For
        End If
        
    Next i
    
    名前列 = Cells.Find("組織スケジュール", LookAt:=xlWhole).Column
    
    For i = 1 To Cells(日付行, Columns.Count).End(xlToLeft).Column
        If TypeName(Cells(日付行, i).Value) = "Date" Then
            月初 = i
            Exit For
        End If
    
    Next i
    
    For i = 月初 To Cells(日付行, Columns.Count).End(xlToLeft).Column
        If Cells(日付行, i).Value = DateSerial(Year(Cells(日付行, 月初).Value), Month(Cells(日付行, 月初).Value) + 1, 0) Then
            月末 = i
            Exit For
        End If
    
    Next i
    
    '---ここから処理範囲を取得する処理---
    
    If Selection(1).Column <> 名前列 Or Selection(1).Row > Cells.Find("組織スケジュール", LookAt:=xlWhole).Row Then
        Set 範囲 = Selection
    Else
        Set 範囲 = Range(Cells(Selection.Row, 月初), Cells(Selection(Selection.Count).Row, 月末))
    End If
    
    '---ここから新規シートを追加する処理---

    Set tmp = ActiveSheet
    Set 転記先シート = Sheets.Add
    tmp.Activate
    転記先シート.Name = Cells(Selection.Row, 名前列).MergeArea(1).Value
    
    '---ここからインポート用のデータを作成する処理---

    i = 1
    
    For Each myRange In 範囲
        If myRange.Value <> "" Then
        
            '---ここから同行者を取得する処理---
            
            同行 = ""
            
            If myRange.Value Like "*休*" Or myRange.Value Like "*ゆとり*" Then
                '休かゆとりが含まれる場合は何もしない
            Else
                For j = 日付行 + 2 To Cells.Find("組織スケジュール", LookAt:=xlWhole).Row - 1
                    If Cells(j, myRange.Column).Value = myRange.Value Then
                        If 同行 = "" Then
                            同行 = Cells(j, 名前列).MergeArea(1).Value
                        Else
                            同行 = 同行 & "、" & Cells(j, 名前列).MergeArea(1).Value
                        End If
                        
                    End If
                    
                Next j
                
                If 同行 = Cells(myRange.Row, 名前列).MergeArea(1).Value Then
                    同行 = ""
                Else
                    同行 = "（" & 同行 & "）"
                End If
                
            End If
            
            '---ここまで---
        
            Set 一致 = Range(Cells(1, 名前列), Cells(Cells(Rows.Count, 名前列).End(xlUp).Row, 名前列)).Find(myRange.Value, LookAt:=xlWhole)
            
            If 一致 Is Nothing Then
                If myRange.Value Like "*" & vbLf & "*" Then
                    myRange.Select
                    MsgBox "処理を中断しました。" & vbCrLf & vbCrLf & myRange.Address(False, False) & "セル" & vbCrLf & "セル内改行"
                    Exit Sub
                End If
                
                予定 = Split(Replace(myRange.Value, "，", ","), ",")
                
                転記先シート.Cells(i, "A").Value = 0
                
                転記先シート.Cells(i, "B").Value = Format(Cells(日付行, myRange.Column).Value, "yyyy/mm/dd")
                
                If UBound(予定) >= 1 Then
                    If IsDate(予定(1)) Then
                        転記先シート.Cells(i, "C").Value = Format(予定(1), "h:mm")
                    Else
                        myRange.Select
                        MsgBox "処理を中断しました。" & vbCrLf & vbCrLf & myRange.Address(False, False) & "セル" & vbCrLf & "時刻として認識できない文字列"
                        Exit Sub
                    End If
                End If
                
                転記先シート.Cells(i, "D").Value = Format(Cells(日付行, myRange.Column).Value, "yyyy/mm/dd")
                
                If UBound(予定) >= 2 Then
                    If IsDate(予定(2)) Then
                        転記先シート.Cells(i, "E").Value = Format(予定(2), "h:mm")
                    Else
                        myRange.Select
                        MsgBox "処理を中断しました。" & vbCrLf & vbCrLf & myRange.Address(False, False) & "セル" & vbCrLf & "時刻として認識できない文字列"
                        Exit Sub
                    End If
                End If
                
                転記先シート.Cells(i, "G").Value = 予定(0) & 同行
                
                myRange.Borders(xlDiagonalUp).LineStyle = True
                myRange.Borders(xlDiagonalDown).LineStyle = True
                
                i = i + 1
                
            Else
                If Cells(一致.Row, myRange.Column).Interior.Color <> 15921906 Then    '塗りつぶしなし（直打ちの内容）の場合
                    If Cells(一致.Row, myRange.Column).Value Like "*" & vbLf & "*" Then
                        Cells(一致.Row, myRange.Column).Select
                        MsgBox "処理を中断しました。" & vbCrLf & vbCrLf & Cells(一致.Row, myRange.Column).Address(False, False) & "セル" & vbCrLf & "セル内改行"
                        Exit Sub
                    End If
                
                    予定 = Split(Replace(Cells(一致.Row, myRange.Column).Value, "，", ","), ",")
                    
                    転記先シート.Cells(i, "A").Value = 0
                    
                    転記先シート.Cells(i, "B").Value = Format(Cells(日付行, myRange.Column).Value, "yyyy/mm/dd")
                    
                    If UBound(予定) >= 1 Then
                        If IsDate(予定(1)) Then
                            転記先シート.Cells(i, "C").Value = Format(予定(1), "h:mm")
                        Else
                            Cells(一致.Row, myRange.Column).Select
                            MsgBox "処理を中断しました。" & vbCrLf & vbCrLf & Cells(一致.Row, myRange.Column).Address(False, False) & "セル" & vbCrLf & "時刻として認識できない文字列"
                            Exit Sub
                        End If
                    End If
                    
                    転記先シート.Cells(i, "D").Value = Format(Cells(日付行, myRange.Column).Value, "yyyy/mm/dd")
                    
                    If UBound(予定) >= 2 Then
                        If IsDate(予定(2)) Then
                            転記先シート.Cells(i, "E").Value = Format(予定(2), "h:mm")
                        Else
                            Cells(一致.Row, myRange.Column).Select
                            MsgBox "処理を中断しました。" & vbCrLf & vbCrLf & Cells(一致.Row, myRange.Column).Address(False, False) & "セル" & vbCrLf & "時刻として認識できない文字列"
                            Exit Sub
                        End If
                    End If
                    
                    転記先シート.Cells(i, "G").Value = myRange.Value & "　" & 予定(0) & 同行
                    
                    myRange.Borders(xlDiagonalUp).LineStyle = True
                    myRange.Borders(xlDiagonalDown).LineStyle = True
                    
                    i = i + 1
                
                Else    '塗りつぶしあり（マクロで転記した内容）の場合
                    転記先シート.Cells(i, "A").Value = 0
                    
                    転記先シート.Cells(i, "B").Value = Format(Cells(日付行, myRange.Column).Value, "yyyy/mm/dd")
                    
                    転記先シート.Cells(i, "D").Value = Format(Cells(日付行, myRange.Column).Value, "yyyy/mm/dd")
                    
                    転記先シート.Cells(i, "G").Value = myRange.Value & "　設備停止対応" & 同行

                    転記先シート.Cells(i, "J").Value = Cells(一致.Row, myRange.Column).Value
                    
                    myRange.Borders(xlDiagonalUp).LineStyle = True
                    myRange.Borders(xlDiagonalDown).LineStyle = True
                    
                    i = i + 1
                    
                End If
                
            End If
            
        End If
        
    Next myRange

End Sub