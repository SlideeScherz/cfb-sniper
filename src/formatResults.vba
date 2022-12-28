Sub formatResults()

  ' Define Range and conditions
  Dim textRng As Range

  ' number of columns for the tabular data
  Dim colNum as Long
  ' hack: not dynamic
  colNum = 108

  ' Setting the STARTING range on which conditional formatting is to be desired
  Set textRng = Worksheets("view").Range("D2:D43")

 ' Iterate through all the rows in the tabular data range
  For colItr = 1 To colNum Step 2 

    ' To delete/clear any existing conditional formatting from the range
    textRng.FormatConditions.Delete

    ' number of rows for the tabular data
    Dim rowNum As Long
    
    ' Capture the number of rows within the tabular data range
    rowNum = textRng.Rows.Count
    
    ' Iterate through all the rows in the tabular data range
    For rowItr = 1 To rowNum
      ' Use a Select Case statement to evaluate the formatting based on column 2
      Select Case textRng.Cells(rowItr, colItr + 1).Value
        Case "TBD."
          textRng.Cells(rowItr, colItr).Font.Color = RGB(0,0,0)
          textRng.Cells(rowItr, colItr).Interior.Color = RGB(200,200,200)
        Case 0
          textRng.Cells(rowItr, colItr).Interior.Color = RGB(255,0,0)
        Case > 0
          textRng.Cells(rowItr, colItr).Interior.Color = RGB(146,208,80)
      End Select
    Next rowItr

  Next colItr
End Sub