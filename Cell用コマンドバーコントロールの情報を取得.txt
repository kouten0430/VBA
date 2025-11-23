  Sub Cell用コマンドバーコントロールの情報を取得()
      Dim i As Integer
      
      For i = 1 To Application.CommandBars("Cell").Controls.Count
          'インデックス番号の取得
          Cells(i, 1) = Application.CommandBars("Cell").Controls(i).Index
          'Cell用コマンドバーコントロールの名前の取得
          Cells(i, 2) = Application.CommandBars("Cell").Controls(i).Caption
          'Cell用コマンドバーコントロールの種類の取得
          Select Case Application.CommandBars("Cell").Controls(i).Type
          Case 1
              Cells(i, 3) = "msoControlButton"
          Case 10
              Cells(i, 3) = "msoControlPopup"
          End Select
      Next
  End Sub