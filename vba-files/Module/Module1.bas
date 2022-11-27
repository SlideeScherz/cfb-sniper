Attribute VB_Name = "Module1"

'/*
' note: Required module
'Add Reference (Tools > References) to the following libraries:
' 1) Microsoft Internet Controls
' 2) Microsoft HTML Object Library
' If edge is not downloaded, download it when you first launch (maybe?)
' https://oxylabs.io/blog/web-scraping-excel-vba
'*/

Sub fetchScores()
  'Add Reference (Tools > References) to the following libraries:
  ' 1) Microsoft Internet Controls
  ' 2) Microsoft HTML Object Library
  
  Dim ie As InternetExplorer
  Dim pagePiece As Object
  Dim webpage As HTMLDocument

  Const SCORES_URL = "https://www.cbssports.com/college-football/scoreboard/"
  
  Set ie = New InternetExplorer
  ie.Visible = True 'Optional if you want to make the window visible
  
  ie.navigate (SCORES_URL)
  Do While ie.readyState = 4: DoEvents: Loop
  Do Until ie.readyState = 4: DoEvents: Loop
  While ie.Busy
    DoEvents
  Wend
  
  Set webpage = ie.document
  Set mtbl = webpage.getElementsByTagName("Table")(1)
  Set table_data = mtbl.getElementsByTagName("tr")
  
  On Error GoTo tryagain:
  For itemNum = 1 To 240
    For childNum = 0 To 5
      Cells(itemNum, childNum + 1) = table_data.Item(itemNum).Children(childNum).innerText
    Next childNum
  Next itemNum
  
  ie.Quit
  Set ie = Nothing
  Exit Sub

  ' routine for error
  tryagain:
    Application.Wait Now + TimeValue("00:00:02")
    errcount = errcount + 1
    Debug.Print Err.Number & Err.Description
    If errcount = 5 Then
      MsgBox "We've detected " & errcount & " errors and we're going to pause the program so you can investigate.", , "Multiple errors detected"
      Stop
      errcount = 0
    End If
    Err.Clear
  Resume
End Sub
