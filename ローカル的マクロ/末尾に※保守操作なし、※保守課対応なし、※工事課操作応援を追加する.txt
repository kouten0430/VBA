Sub 末尾に※保守操作なしを追加する()
    'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "※保守操作なし"
        
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルにのみ処理を行う
            If myRange.Value <> "" Then 'セルが空白でなければ改行し、空白であれば改行しない
                myRange.Value = myRange.Value & vbLf & V
            Else
                myRange.Value = V
            End If
            
                    '---ここから指定した文字の色を変える処理---

                    'myRange（Rangeオブジェクト）、V（指定文字）、255（文字色）を必要に応じて変更して下さい
                    '指定文字が複数ある場合はInStrでヒットする最初の文字列のみ色が変わります

                    myRange.Characters(InStr(myRange.Value, V), Len(V)).Font.Color = 255

                    '---ここまで---

        End If
        
    Next myRange
    
End Sub
Sub 末尾に※保守課対応なしを追加する()
    'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "※保守課対応なし"
        
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルにのみ処理を行う
            If myRange.Value <> "" Then 'セルが空白でなければ改行し、空白であれば改行しない
                myRange.Value = myRange.Value & vbLf & V
            Else
                myRange.Value = V
            End If
            
                    '---ここから指定した文字の色を変える処理---

                    'myRange（Rangeオブジェクト）、V（指定文字）、255（文字色）を必要に応じて変更して下さい
                    '指定文字が複数ある場合はInStrでヒットする最初の文字列のみ色が変わります

                    myRange.Characters(InStr(myRange.Value, V), Len(V)).Font.Color = 255

                    '---ここまで---

        End If
        
    Next myRange
    
End Sub
Sub 末尾に※工事課操作応援を追加する()
    'ローカル的なマクロです
    Dim V As String
    Dim myRange As Range
    
    V = "※工事課操作応援"
        
    For Each myRange In Selection
        If myRange.Address = myRange.MergeArea(1).Address Then   '結合セルの場合は左上のセルにのみ処理を行う
            If myRange.Value <> "" Then 'セルが空白でなければ改行し、空白であれば改行しない
                myRange.Value = myRange.Value & vbLf & V
            Else
                myRange.Value = V
            End If
            
                    '---ここから指定した文字の色を変える処理---

                    'myRange（Rangeオブジェクト）、V（指定文字）、255（文字色）を必要に応じて変更して下さい
                    '指定文字が複数ある場合はInStrでヒットする最初の文字列のみ色が変わります

                    myRange.Characters(InStr(myRange.Value, V), Len(V)).Font.Color = 255

                    '---ここまで---

        End If
        
    Next myRange
    
End Sub