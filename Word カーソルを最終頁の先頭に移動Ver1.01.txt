Sub カーソルを最終頁の先頭に移動()
    Dim 頁 As Integer
    
    'ActiveWindow.ActivePane.View.Zoom.Percentage = 74
    ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitFullPage
    
    頁 = Selection.Information(wdNumberOfPagesInDocument)
    
    Selection.GoTo What:=wdGoToPage, Count:=頁, Which:=wdGoToFirst
    
    Selection.MoveDown Unit:=wdScreen, Count:=1
    
End Sub