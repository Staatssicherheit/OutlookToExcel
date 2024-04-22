# Load Selenium WebDriver module
Add-Type -Path "path\to\Selenium.WebDriver.dll"
Add-Type -Path "path\to\Selenium.WebDriver.Support.dll"

# Create an instance of Excel
$excel = New-Object -ComObject Excel.Application

# Make Excel visible (optional)
$excel.Visible = $true

# Add a new workbook
$workbook = $excel.Workbooks.Add()

# Add a new worksheet
$worksheet = $workbook.Worksheets.Add()

# Set up Selenium WebDriver
$driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver

# Navigate to the login page
$driver.Navigate().GoToUrl("https://example.com/login")

# Find username and password fields and login button
$usernameField = $driver.FindElementByXPath("//input[@id='username']")
$passwordField = $driver.FindElementByXPath("//input[@id='password']")
$loginButton = $driver.FindElementByXPath("//button[@id='loginButton']")

# Input username and password
$usernameField.SendKeys("your_username")
$passwordField.SendKeys("your_password")

# Click the login button
$loginButton.Click()

# Wait for the page to load
Start-Sleep -Seconds 5

# Get the HTML content of the page
$html = $driver.PageSource

# Paste HTML content into Excel
$worksheet.Cells.Item(1,1).Value2 = $html

# Save Excel workbook (optional)
$workbook.SaveAs("path\to\output.xlsx")

# Close the WebDriver
$driver.Quit()

# Close Excel application
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
