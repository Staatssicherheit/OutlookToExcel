Sub GRP01_03_05()
Dim driver As WebDriver

Dim SBCURLS As Variant
Dim myID As Variant
Dim myPW As Variant
Dim Data As Object

SBCURLS = Array("http://", "http://", "http://")
myID = Array("ID", "ID", "ID")
myPW = Array("PW", "PW", "PW")

Dim i As Integer
For i = LBound(SBCURLS) To UBound(SBCURLS)

Set driver = New WebDriver

driver.Start "Chrome"
driver.Wait (3000)
driver.Get SBCURLS(i)

driver.FindElementById("ws_loginname").SendKeys myID(i)
driver.FindElementById("ws_loginpass").SendKeys myPW(i)
driver.FindElementById("login_button").Click

driver.Wait (3000)

Set Data = driver.FindElementByCss("#trunkTBL > table")

Dim xlApp As Object
Set xlApp = GetObject(, "Excel.Application")
Dim xlBook As Object
Set xlBook = xlApp.ActiveWorkbook
Dim xlSheet As Object
Set xlSheet = xlBook.Sheets.Add

Dim rowNum As Integer
Dim colNum As Integer
rowNum = 1
For Each Row In Data.FindElementsByTag("tr")
    colNum = 1
    For Each Cell In Row.FindElementsByTag("td")
        xlSheet.Cells(rowNum, colNum).Value = Cell.Text
        colNum = colNum + 1
    Next Cell
    rowNum = rowNum + 1
Next Row

Select Case i
    Case 0
        xlSheet.Name = "M4K_GRP01"
    Case 1
        xlSheet.Name = "M4K_GRP03"
    Case 2
        xlSheet.Name = "M4K_GRP05"
End Select

driver.Wait (3000)

driver.Close

Set driver = Nothing
Next i
End Sub
