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

  Set numTables = webpage.getElementsByClassName("score-cards") 
  Debug.Print "Tables found:"
  Debug.Print numTables.children().length()

  Set mtbl = webpage.getElementsByTagName("table")(2) ' loop thru this for x many tables
  Set table_data = mtbl.getElementsByTagName("tr")
  Debug.Print "fetched", table_data.length(), "elements"
  
  ' NOTE: this is all for one game.
  ' team name, q1,q2,q3,q4,total (5 elements)
  Cells(1, 1) = table_data.Item(0).Children(0).innerText
  Cells(1, 2) = table_data.Item(0).Children(1).innerText
  Cells(1, 3) = table_data.Item(0).Children(2).innerText
  Cells(1, 4) = table_data.Item(0).Children(3).innerText
  Cells(1, 5) = table_data.Item(0).Children(4).innerText
  Cells(1, 6) = table_data.Item(0).Children(5).innerText

  Cells(2, 1) = table_data.Item(1).Children(0).innerText
  Cells(2, 2) = table_data.Item(1).Children(1).innerText
  Cells(2, 3) = table_data.Item(1).Children(2).innerText
  Cells(2, 4) = table_data.Item(1).Children(3).innerText
  Cells(2, 5) = table_data.Item(1).Children(4).innerText
  Cells(2, 6) = table_data.Item(1).Children(5).innerText

  Cells(3, 1) = table_data.Item(2).Children(0).innerText
  Cells(3, 2) = table_data.Item(2).Children(1).innerText
  Cells(3, 3) = table_data.Item(2).Children(2).innerText
  Cells(3, 4) = table_data.Item(2).Children(3).innerText
  Cells(3, 5) = table_data.Item(2).Children(4).innerText
  Cells(3, 6) = table_data.Item(2).Children(5).innerText
  
  'Cells(itemNum, childNum + 1).Interior.Color = RGB(246, 174, 134)
  
  ie.Quit
  Set ie = Nothing

End Sub
