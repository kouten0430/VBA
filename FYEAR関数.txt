Function FYEAR(KYEAR As Date) As Integer
    '引数の日付データから年度を取得します
    If Month(KYEAR) <= 3 Then
        FYEAR = Year(KYEAR) - 1
    Else
        FYEAR = Year(KYEAR)
    End If
End Function