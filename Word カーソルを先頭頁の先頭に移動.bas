Sub カーソルを先頭頁の先頭に移動()
    Dim 頁 As Integer
    
    'ActiveWindow.ActivePane.View.Zoom.Percentage = 74
    ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitFullPage
    
    頁 = 1
    
    Selection.GoTo What:=wdGoToPage, Count:=頁, Which:=wdGoToFirst
    
End Sub