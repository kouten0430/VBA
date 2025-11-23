Sub フレームも含めてテーブルを全部書き出す()
    Dim ie As Object
    Dim sh As Object
    Dim win As Object
    Dim i As Long
    Dim r As Long
    Dim c As Integer
    Dim j As Integer
    Dim k As Long
    
    Set sh = CreateObject("Shell.Application")
    
    For Each win In sh.Windows
        If win.Name = "Internet Explorer" Then
            Set ie = win
            Exit For
        End If
    Next
    
    r = 1
    c = 1
    
    For i = 0 To ie.document.all.Length - 1
        If ie.document.all(i).tagName = "TH" Or ie.document.all(i).tagName = "TD" Then
            Cells(r, c) = ie.document.all(i).innerText
            c = c + 1
        ElseIf ie.document.all(i).tagName = "TR" Then
            r = r + 1
            c = 1
        End If
    Next i
    
    For j = 0 To ie.document.frames.Length - 1
        For k = 0 To ie.document.frames(j).document.all.Length - 1
            If ie.document.frames(j).document.all(k).tagName = "TH" Or ie.document.frames(j).document.all(k).tagName = "TD" Then
                Cells(r, c) = ie.document.frames(j).document.all(k).innerText
                c = c + 1
            ElseIf ie.document.frames(j).document.all(k).tagName = "TR" Then
                r = r + 1
                c = 1
            End If
        Next k
    Next j
        
End Sub