Sub 選択位置からデータ群の右端まで列全体を選択()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    If Selection.Count > 1 Then Exit Sub
    Range(Cells(ActiveCell.Row, ActiveCell.Column), Cells(ActiveCell.Row, ActiveCell.CurrentRegion.Columns(ActiveCell.CurrentRegion.Columns.Count).Column)).EntireColumn.Select
End Sub
Sub 選択位置からデータ群の左端まで列全体を選択()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    If Selection.Count > 1 Then Exit Sub
    Range(Cells(ActiveCell.Row, ActiveCell.CurrentRegion.Column), Cells(ActiveCell.Row, ActiveCell.Column)).EntireColumn.Select
End Sub
Sub 選択位置からデータ群の下端まで行全体を選択()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    If Selection.Count > 1 Then Exit Sub
    Range(Cells(ActiveCell.Row, ActiveCell.Column), Cells(ActiveCell.CurrentRegion.Rows(ActiveCell.CurrentRegion.Rows.Count).Row, ActiveCell.Column)).EntireRow.Select
End Sub
Sub 選択位置からデータ群の上端まで行全体を選択()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    If Selection.Count > 1 Then Exit Sub
    Range(Cells(ActiveCell.CurrentRegion.Row, ActiveCell.Column), Cells(ActiveCell.Row, ActiveCell.Column)).EntireRow.Select
End Sub
Sub 選択範囲の右側を拡張()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + 1).Select
End Sub
Sub 選択範囲の右側を縮小()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count - 1).Select
End Sub
Sub 選択範囲を下側を拡張()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    Selection.Resize(Selection.Rows.Count + 1, Selection.Columns.Count).Select
End Sub
Sub 選択範囲を下側を縮小()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    Selection.Resize(Selection.Rows.Count - 1, Selection.Columns.Count).Select
End Sub
Sub 選択範囲の左側を拡張()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    Selection.Offset(0, -1).Resize(Selection.Rows.Count, Selection.Columns.Count + 1).Select
End Sub
Sub 選択範囲の左側を縮小()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    Selection.Offset(0, 1).Resize(Selection.Rows.Count, Selection.Columns.Count - 1).Select
End Sub
Sub 選択範囲の上側を拡張()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    Selection.Offset(-1, 0).Resize(Selection.Rows.Count + 1, Selection.Columns.Count).Select
End Sub
Sub 選択範囲の上側を縮小()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1, Selection.Columns.Count).Select
End Sub
Sub 選択位置からデータが連続する右端まで列全体を選択()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    If Selection.Count > 1 Then Exit Sub
    Range(Cells(ActiveCell.Row, ActiveCell.Column), Cells(ActiveCell.Row, ActiveCell.End(xlToRight).Column)).EntireColumn.Select
End Sub
Sub 選択位置からデータが連続する左端まで列全体を選択()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    If Selection.Count > 1 Then Exit Sub
    Range(Cells(ActiveCell.Row, ActiveCell.End(xlToLeft).Column), Cells(ActiveCell.Row, ActiveCell.Column)).EntireColumn.Select
End Sub
Sub 選択位置からデータが連続する下端まで行全体を選択()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    If Selection.Count > 1 Then Exit Sub
    Range(Cells(ActiveCell.Row, ActiveCell.Column), Cells(ActiveCell.End(xlDown).Row, ActiveCell.Column)).EntireRow.Select
End Sub
Sub 選択位置からデータが連続する上端まで行全体を選択()
    Application.ScreenUpdating = False  '画面表示の更新をオフにする
    If Selection.Count > 1 Then Exit Sub
    Range(Cells(ActiveCell.End(xlUp).Row, ActiveCell.Column), Cells(ActiveCell.Row, ActiveCell.Column)).EntireRow.Select
End Sub