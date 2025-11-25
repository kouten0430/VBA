  Sub GetCommandbarInfo()
      Dim AppCmdBar As CommandBar
      Dim i As Integer
      i = 0
      For Each AppCmdBar In Application.CommandBars
          i = i + 1
          'インデックス番号の取得
          Cells(i, 1) = AppCmdBar.Index
          'コマンドバーの名前の取得
          Cells(i, 2) = AppCmdBar.Name
          'コマンドバーの種類の取得
          Select Case AppCmdBar.Type
          Case 0
              Cells(i, 3) = "msoBarTypeNomal"
          Case 1
              Cells(i, 3) = "msoBarTypeMenuBar"
          Case 2
              Cells(i, 3) = "msoBarTypePopup"
          End Select
      Next AppCmdBar
  End Sub