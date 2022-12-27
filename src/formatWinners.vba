
Sub formatWinners()

  ' Define Range and conditions
  Dim textRng As Range
  'Dim winCondition As FormatCondition, lossCondition As FormatCondition

  'Fixing/Setting the range on which conditional formatting is to be desired
  Set textRng = Worksheets("view").Range("D2:D43")

  'To delete/clear any existing conditional formatting from the range
  textRng.FormatConditions.Delete

  'Defining and setting the criteria for each conditional format
  'Set winCondition = textRng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
  'Set lossCondition = textRng.FormatConditions.Add(xlCellValue, xlEqual, "=0")

    'Defining and setting the format to be applied for each condition
'  With winCondition
 '   .Font.Color = vbGreen
  '  .Font.Bold = False
  'End With

  'With lossCondition
   ' .Font.Color = vbRed
    '.Font.Bold = False
  'End With

  ' itrs
  Dim i As Long
  Dim c As Long
  Dim testcell As Range
  c = textRng.Cells.Count

  ' loop
  For i = 1 To c
    Set testcell = textRng(i,2)
    Select Case testcell
      Case Is > 0
        With testcell
          .Interior.Color = RGB(0,128,0)
          '.Font.Color = vbWhite
        End With
      Case Is = 0
        With testcell
          .Interior.Color = RGB(200,0,0)
          .Font.Color = vbWhite
        End With
    End Select
  Next i

End Sub
