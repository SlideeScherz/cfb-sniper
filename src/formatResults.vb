Sub formatResults()

  ' Define Range and conditions
  Dim textRng As Range
  Dim winCondition As FormatCondition, lossCondition As FormatCondition
  
  'Fixing/Setting the range on which conditional formatting is to be desired
  Set textRng = Worksheets("view").Range("D2:D43")

  'To delete/clear any existing conditional formatting from the range
  textRng.FormatConditions.Delete

  'Create variables to hold the number of rows for the tabular data
  Dim RRow As Long, N As Long

  'Capture the number of rows within the tabular data range
  RRow = textRng.Rows.Count
  
  'Iterate through all the rows in the tabular data range
  For N = 1 To RRow
    'Use a Select Case statement to evaluate the formatting based on column 2
    Select Case textRng.Cells(N, 2).Value
      Case "TBD."
        textRng.Cells(N, 1).Font.Color = vbGray
      Case 0
        textRng.Cells(N, 1).Interior.Color = RGB(255,0,0)
      Case > 0
        textRng.Cells(N, 1).Interior.Color = RGB(146,208,80)
    End Select
  Next N
End Sub