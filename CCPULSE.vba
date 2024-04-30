Sub CopyTableFromApplicationToExcel()
    Dim appPath As String
    Dim appHandle As Long
    Dim xlApp As Object
    Dim xlSheet As Object
    
    ' Specify the path of the application
    appPath = "C:\Path\To\Your\Application.exe"
    
    ' Open the application
    appHandle = Shell(appPath, vbNormalFocus)
    
    ' Wait for the application to open
    Application.Wait Now + TimeValue("0:00:05") ' Adjust the wait time as needed
    
    ' Send keys for login (replace with your login details)
    SendKeys "username", True
    SendKeys "{TAB}", True
    SendKeys "password", True
    SendKeys "{ENTER}", True
    
    ' Wait for the login process to complete
    Application.Wait Now + TimeValue("0:00:05") ' Adjust the wait time as needed
    
    ' Send keys to navigate to the table (replace with your navigation keys)
    SendKeys "{TAB}", True
    SendKeys "{ENTER}", True
    
    ' Wait for the table to be displayed
    Application.Wait Now + TimeValue("0:00:05") ' Adjust the wait time as needed
    
    ' Select all and copy the table
    SendKeys "^a", True ' Select all
    SendKeys "^c", True ' Copy
    
    ' Create a new instance of Excel
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = True ' Make Excel visible
    
    ' Add a new workbook
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    
    ' Paste the copied table into Excel
    xlSheet.Range("A1").Select ' Select the desired starting cell
    xlSheet.PasteSpecial Format:="Text", Link:=False, DisplayAsIcon:=False
    
    ' Save the Excel file
    xlBook.SaveAs "C:\Path\To\Your\Excel\File.xlsx" ' Specify the file path
    
    ' Close the Excel workbook and application
    xlBook.Close
    xlApp.Quit
    
    ' Release objects
    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub
