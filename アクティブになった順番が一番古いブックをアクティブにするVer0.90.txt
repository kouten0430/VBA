Sub アクティブになった順番が一番古いブックをアクティブにする()
    '現在アクティブなブックのインデックスが１，その前が２，その前が３……と変わるためループする
    Windows(Windows.Count - 1).Activate
    
End Sub