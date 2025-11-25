Sub 出退社記録日直()
    Dim ie As InternetExplorer
    Dim htdoc As HTMLDocument
    Dim frame1 As HTMLDocument
    Dim span1 As HTMLSpanElement
    Dim span2 As HTMLSpanElement
    Dim input1 As HTMLInputElement
    Dim input2 As HTMLInputElement
    Dim input3 As HTMLInputElement
    Dim input4 As HTMLInputElement
    Dim input5 As HTMLInputElement
    Dim input6 As HTMLInputElement
    Dim select1 As HTMLSelectElement
    Dim 出社時刻 As Date
    Dim 退社時刻 As Date
    Dim 出社実績 As Date
    Dim 退社実績 As Date
    Dim 出社予定 As Date
    Dim 退社予定 As Date

    Set ie = getIE("労働時間一元管理システム")
    
    If ie Is Nothing Then
        MsgBox "対象画面が見つかりません"
        Exit Sub
    End If
   
    Set htdoc = ie.document
    Set frame1 = htdoc.frames("contents").document
    Set span1 = frame1.getElementById("lblSICJikoku") '出社時刻
    Set span2 = frame1.getElementById("lblEICJikoku") '退社時刻
    Set input1 = frame1.getElementById("txtSJiSyotei")  '出社実績
    Set input2 = frame1.getElementById("txtEJiSyotei")  '退社実績
    Set input3 = frame1.getElementById("txtSYoSyotei")  '出社予定
    Set input4 = frame1.getElementById("txtEYoSyotei")  '退社予定
    Set input5 = frame1.getElementById("btnAutoCreate")  '除外時間自動生成
    Set input6 = frame1.getElementById("txtSGyoumuNaiyou")  '業務内容（実績）
    Set select1 = frame1.getElementById("ddlEYoSijiCd1")    '所属長指示
        
    出社時刻 = span1.innerText
    退社時刻 = span2.innerText

    出社実績 = WorksheetFunction.Ceiling(出社時刻, 10 / (24 * 60))  '10分単位で切り上げ
    出社実績 = WorksheetFunction.Ceiling(出社実績 + TimeValue("0:05"), 10 / (24 * 60))  '必ず10～19分の差が付くようにする
    
    If 出社時刻 <= #9:40:00 AM# And 出社時刻 >= #9:20:00 AM# Then
        出社実績 = #9:40:00 AM#
        出社予定 = #9:40:00 AM#
    ElseIf 出社時刻 < #6:55:00 AM# Then
        出社実績 = #7:00:00 AM#
        出社予定 = #7:00:00 AM#
    Else
        出社予定 = 出社実績
    End If

    退社実績 = WorksheetFunction.Floor(退社時刻, 10 / (24 * 60))   '10分単位で切り捨て
    退社実績 = WorksheetFunction.Floor(退社実績 - TimeValue("0:05"), 10 / (24 * 60))    '必ず10～19分の差が付くようにする
    
    If 退社時刻 <= #6:40:00 PM# And 退社時刻 >= #6:20:00 PM# Then
        退社実績 = #6:20:00 PM#
        退社予定 = #6:20:00 PM#
    ElseIf 退社時刻 > #9:05:00 PM# Then
        退社実績 = #9:00:00 PM#
        退社予定 = #9:00:00 PM#
    ElseIf 退社時刻 < #6:20:00 PM# Then
        退社予定 = 退社実績
    Else
        退社予定 = WorksheetFunction.Ceiling(退社実績, 30 / (24 * 60))  '30分単位で切り上げ
    End If

    input1.Value = Format(出社実績, "hhmm")
    input2.Value = Format(退社実績, "hhmm")
    input3.Value = Format(出社予定, "hhmm")
    input4.Value = Format(退社予定, "hhmm")
    input6.Value = "日直"
    select1.Value = "0001"
    input5.Click
    
End Sub