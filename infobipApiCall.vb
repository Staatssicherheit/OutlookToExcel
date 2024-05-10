Sub GetSmsDeliveryRate()
    ' Set the API endpoint for SMS delivery reports
    Dim apiUrl As String
    apiUrl = "https://api.infobip.com/sms/1/reports"

    ' Set your InfoBip API credentials (username and password)
    Dim username As String
    Dim password As String
    username = "your_username"
    password = "your_password"

    ' Set the phone number and message content
    Dim phoneNumber As String
    Dim message As String
    phoneNumber = "41793026727"
    message = "This is a sample message"

    ' Make the API call
    Dim response As String
    response = GetApiResponse(apiUrl, username, password, phoneNumber, message)

    ' Parse the JSON response and extract the delivery rate
    Dim json As Object
    Set json = JsonConverter.ParseJson(response)
    Dim deliveryRate As Double
    deliveryRate = json("results")(1)("status")("deliveryRate")

    ' Print the delivery rate to the Immediate Window
    Debug.Print "SMS Delivery Rate: " & deliveryRate
End Sub

Function GetApiResponse(apiUrl As String, username As String, password As String, phoneNumber As String, message As String) As String
    ' Create an HTTP request object
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Construct the API URL with query parameters
    Dim fullUrl As String
    fullUrl = apiUrl & "?username=" & username & "&password=" & password & "&to=" & phoneNumber & "&text=" & message

    ' Open a connection to the API URL
    xmlHttp.Open "GET", fullUrl, False
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    xmlHttp.send ""

    ' Get the API response
    GetApiResponse = xmlHttp.responseText
End Function
