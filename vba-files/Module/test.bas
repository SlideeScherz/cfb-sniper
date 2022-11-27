' tools > references >
' Microsoft HTML Client Library and Microsoft Internet Control
' must be enabled

' If edge is not downloaded, download it when you first launch (maybe?)
' https://oxylabs.io/blog/web-scraping-excel-vba

Sub espnScraper()
  Dim browser As InternetExplorer
  Dim page As HTMLDocument
  Dim quotes As Object
  Dim authors As Object

  Set browser = New InternetExplorer
  browser.Visible = True
  browser.Navigate ("https://www.cbssports.com/college-football/scoreboard/")
  Do While browser.Busy: Loop

  Set page = browser.document
  ' team team--collegefootball
  Set quotes = page.getElementsByClassName("quote")
  Set authors = page.getElementsByClassName("author")

  For num = 1 To 5
    Cells(num, 1).Value = quotes.Item(num).innerText
    Cells(num, 2).Value = authors.Item(num).innerText
  Next num

  browser.Quit
    
End Sub
