Sub GitHubAutomation()
    Dim bot As New Selenium.ChromeDriver
    bot.Start "chrome", "https://github.com/login"
    
    bot.Get "/login"
    bot.FindElementByName("login").SendKeys "YourUsername"
    bot.FindElementByName("password").SendKeys "YourPassword"
    bot.FindElementByName("commit").Click
    
    ' Wait for page to load after login
    bot.Wait 5000 ' Adjust as needed
    
    bot.Get "URL_of_the_page_with_table_data"
    
    ' Extract table data and copy to Excel
    Dim tableData As Object
    Set tableData = bot.FindElementByXPath("//table")
    
    ' Copy table data to Excel
    ' Example: Range("A1").Value = tableData.Text
    
    bot.Quit
End Sub
