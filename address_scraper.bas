Attribute VB_Name = "address_scraper"
Option Explicit

Sub GrabSP500Address()
    
    Dim timeStart As Double
    timeStart = Timer
    
    Application.ScreenUpdating = False
    
    'SEC EDGAR URL
    Dim secURL As String
    secURL = "https://www.sec.gov/cgi-bin/browse-edgar?CIK="
    
    Dim queryString As String
    queryString = "&action=getcompany"
    
'    Needs Microsoft HTML Object Library and Microsoft Internet Controls References Enabled
'    Otherwise, use the commented code below.
'    Dim IE As Object
'    Set IE = CreateObject("InternetExplorer.Application")
    
    Dim IE As InternetExplorer
    Set IE = New InternetExplorer
    IE.Visible = False
    
    Dim objCollection As Object
    Dim objElement As Object
    
    Dim i As Integer
    Dim colOffset As Integer
    
    Dim cik As Range
    'Company CIKs are in column H. Could use stock ticker symbol in column A if necessary.
    'Ticker symbols won't work with companies that have a period in their names.
    For Each cik In Range("H2:H506")
        'Navigate to page and wait till ready.
        IE.Navigate secURL & cik.Value & queryString
        Do While IE.Busy Or IE.readyState <> 4 '4 = READYSTATE_COMPLETE
            DoEvents
        Loop
        
        Set objCollection = IE.document.getElementsByTagName("span")
        i = 0
        colOffset = 0
        
        'Loop through span tags and look for class = mailerAddress.
        'Write information between span tags in Excel.
        Do While i < objCollection.Length
            If objCollection(i).className = "mailerAddress" Then
                Cells(cik.Row, "I").Offset(0, colOffset).Value = objCollection(i).innerText
                colOffset = colOffset + 1
            End If
            i = i + 1
        Loop
    Next
    
    'Clean up.
    IE.Quit
    Set IE = Nothing
    
    Application.ScreenUpdating = True
    
    Dim timeEnd As Double
    timeEnd = Round(Timer - timeStart, 2)
    
    MsgBox "Code ran successfully in " & timeEnd & " seconds."
    
End Sub
