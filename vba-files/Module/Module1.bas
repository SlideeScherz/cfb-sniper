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
  Dim mtbl, table_data As Object
  Dim webpage As HTMLDocument

  Const SCORES_URL = "https://www.cbssports.com/college-football/scoreboard/"
  
  Set ie = New InternetExplorer
  'ie.Visible = True 'Optional if you want to make the window visible
  
  ie.navigate (SCORES_URL)

  Do
  DoEvents
  ' await webpage load (loop through nothing until ready)
  Loop Until ie.readyState = READYSTATE_COMPLETE

  Debug.Print "Browser ready."

  Set webpage = ie.document
  Set mtbl = webpage.getElementsByTagName("table")
  Set table_data = mtbl.getElementsByTagName("tr")
  
  Dim numGames As Integer
  numGames = Len(mtbl)
  
  

  Debug.Print "fetched"; numGames; "elements"
  
  For Count = 0 To numGames
    Debug.Print mtbl.Item(Count).Children(0).innerText
  Next Count   
  
  
  
  ' NOTE: this is all for one game.
  ' team name, q1,q2,q3,q4,total (5 elements)
  'Cells(1, 1) = table_data.Item(0).Children(0).innerText

  'Cells(1, 1) = table_data.Item(0).Children(0).innerText
  
  'Cells(itemNum, childNum + 1).Interior.Color = RGB(246, 174, 134)
  
  ie.Quit
  Set ie = Nothing

End Sub
