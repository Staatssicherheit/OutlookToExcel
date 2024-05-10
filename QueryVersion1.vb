Sub SendRequestUsingWinHttp()
    Dim winHttpRequest As Object
    Set winHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Specify the URL of the server to which you want to send a request
    Dim url As String
    url = "Your Infobip Metrics API Endpoint"
    
    ' Open the HTTP request
    winHttpRequest.Open "POST", url, False
    
    ' Set any necessary request headers
    winHttpRequest.setRequestHeader "Content-Type", "application/json"
    winHttpRequest.setRequestHeader "Authorization", "Basic YourBase64EncodedCredentials"
    
    ' Send the request with the necessary parameters
    Dim jsonBody As String
    jsonBody = "Your JSON Payload"
    winHttpRequest.Send jsonBody
    
    ' Handle the response
    Dim responseText As String
    responseText = winHttpRequest.responseText
    
    ' Process the responseText as needed
End Sub
