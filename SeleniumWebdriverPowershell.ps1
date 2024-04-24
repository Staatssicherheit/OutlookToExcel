# Set the path to the directory containing ChromeDriver
$chromeDriverPath = "C:\Path\To\ChromeDriver"

# Set the URL of the webpage you want to automate
$url = "https://example.com/login"

# Set the path where you want to save the Excel file
$excelFilePath = "C:\Path\To\Save\ExcelFile.xlsx"

# Start ChromeDriver
$driver = Start-SeChrome -Path $chromeDriverPath

# Navigate to the login page
$driver.Navigate().GoToUrl($url)

# Find the username and password input fields and enter your credentials
$usernameField = $driver.FindElementByName("username")
$usernameField.SendKeys("your_username")
$passwordField = $driver.FindElementByName("password")
$passwordField.SendKeys("your_password")

# Find and click the login button
$loginButton = $driver.FindElementById("login-button")
$loginButton.Click()

# Wait for the page to load after login
Start-Sleep -Seconds 5

# Navigate to the table view page
$tableViewLink = $driver.FindElementByXPath("//a[contains(@href, 'table-view')]")
$tableViewLink.Click()

# Wait for the table view page to load
Start-Sleep -Seconds 5

# Find the table element and extract its data
$table = $driver.FindElementByXPath("//table[@id='table-id']")
$tableData = $table.Text

# Write the data to an Excel file
$tableData | Out-File -FilePath $excelFilePath

# Close the ChromeDriver
$driver.Quit()
