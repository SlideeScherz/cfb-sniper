
Sub formatWinners()

  ' Define Range and conditions
  Dim textRng As Range
  Dim winCondition As FormatCondition, lossCondition As FormatCondition
  
  'Fixing/Setting the range on which conditional formatting is to be desired
  Set textRng = Worksheets("view").Range("E2:E43")

  'To delete/clear any existing conditional formatting from the range
  textRng.FormatConditions.Delete

  'Defining and setting the criteria for each conditional format
  Set lossCondition = textRng.FormatConditions.Add(xlCellValue, xlEqual, "=0")
  Set winCondition = textRng.FormatConditions.Add(xlCellValue, xlGreater, "=0")

  With winCondition
    .Interior.Color = RGB(0, 150, 0)
    '.Font.Color = RGB(0,128,0)
    .Font.Bold = False
  End With

  With lossCondition
    .Interior.Color = RGB(200, 0, 0)
    '.Font.Color = RGB(150,0,0)
    .Font.Bold = False
  End With

End Sub

