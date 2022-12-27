
Sub formatWinners()

  ' Define Range and conditions
  Dim textRng As Range
  Dim winCondition As FormatCondition, lossCondition As FormatCondition

  'Fixing/Setting the range on which conditional formatting is to be desired
  Set textRng = Worksheets("view").Range("E2:E43")

  'To delete/clear any existing conditional formatting from the range
  textRng.FormatConditions.Delete

  'Defining and setting the criteria for each conditional format
  Set winCondition = textRng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
  Set lossCondition = textRng.FormatConditions.Add(xlCellValue, xlEqual, "=0")

  'Defining and setting the format to be applied for each condition
  With winCondition
    .Font.Color = vbGreen
    .Font.Bold = False
  End With

  With lossCondition
    .Font.Color = vbRed
    .Font.Bold = False
  End With

End Sub
