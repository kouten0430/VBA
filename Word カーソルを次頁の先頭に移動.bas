Sub カーソルを次頁の先頭に移動()
    'ActiveWindow.ActivePane.View.Zoom.Percentage = 74
    ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitFullPage
    
    Selection.MoveDown Unit:=wdScreen, Count:=1
    
End Sub