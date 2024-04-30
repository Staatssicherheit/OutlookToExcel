Sub doLoginAndGetTable()
    Dim driver As SeleniumVBA.WebDriver
    Dim keys As SeleniumVBA.WebKeyboard
    Dim actions As SeleniumVBA.WebActionChain
    Dim searchBox As SeleniumVBA.WebElement
    
    Set driver = SeleniumVBA.New_WebDriver
    Set keys = SeleniumVBA.New_WebKeyboard
    
    driver.StartChrome
    
    driver.OpenBrowser
    
    driver.NavigateTo "http://182.195.78.170/SASRegUsers"
    driver.Wait 500
    
    loginID = "gccshelp"
    loginPW = "12#$qwER"
    
    driver.FindElement(By.ID, "ws_loginname").SendKeys loginID
    driver.FindElement(By.ID, "ws_loginpass").SendKeys loginPW
    
    driver.FindElement(By.ID, "login_button").Click
    
    driver.Wait 2000
    
    driver.NavigateTo "http://182.195.78.170/SASRegUsers"
    
    driver.Wait 2000
    
    Dim tableData As Object
    Set tableData = driver.FindElement(By.ClassName, "tb")
    
    
    Range("A1").Value = tableData.text
    
    driver.CloseBrowser
    driver.Shutdown
End Sub
