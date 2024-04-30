Sub CopyDataFromWebToExcel()
    Dim IE As Object
    Dim HTMLDoc As Object
    Dim Username As Object
    Dim Password As Object
    Dim LoginButton As Object
    Dim Data As Object
    Dim i As Integer
    
    ' Create a new instance of InternetExplorer
    Set IE = CreateObject("InternetExplorer.Application")
    
    ' Navigate to the website
    IE.Navigate "https://example.com/login"
    
    ' Wait for the webpage to load
    Do While IE.readyState <> 4 Or IE.Busy
        DoEvents
    Loop
    
    ' Get the document object
    Set HTMLDoc = IE.document
    
    ' Fill in the username and password fields
    HTMLDoc.getElementById("username").Value = "YourUsername"
    HTMLDoc.getElementById("password").Value = "YourPassword"
    
    ' Locate and click the login button
    Set LoginButton = HTMLDoc.getElementById("loginButton")
    LoginButton.Click
    
    ' Wait for the page to load after login
    Do While IE.readyState <> 4 Or IE.Busy
        DoEvents
    Loop
    
    ' Extract data from the webpage
    Set Data = HTMLDoc.getElementById("tableID") ' Assuming the data is in a table
    
    ' Copy data to Excel
    For i = 0 To Data.Rows.Length - 1
        For j = 0 To Data.Rows(i).Cells.Length - 1
            ThisWorkbook.Sheets("Sheet1").Cells(i + 1, j + 1).Value = Data.Rows(i).Cells(j).innerText
        Next j
    Next i
    
    ' Close IE
    IE.Quit
    Set IE = Nothing
    
    MsgBox "Data copied to Excel successfully!"
End Sub
