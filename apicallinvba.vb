' Add a reference to Microsoft XML, v6.0 before running this code.
' Go to Tools > References in the VBA editor to add the reference.

Sub GetApiData()
    ' Set your API endpoint URL
    Dim apiUrl As String
    apiUrl = "https://jsonplaceholder.typicode.com/users"

    ' Call the API and get the response
    Dim response As String
    response = GetApiResponse(apiUrl)

    ' Parse JSON response
    Dim json As Object
    Set json = JsonConverter.ParseJson(response)

    ' Process JSON data and populate Excel sheet
    PopulateSheet json
End Sub

Function GetApiResponse(apiUrl As String) As String
    ' Create HTTP request object
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Open a connection to the API URL
    xmlHttp.Open "GET", apiUrl, False
    xmlHttp.setRequestHeader "Content-Type", "application/json"
    xmlHttp.send ""

    ' Get the API response
    GetApiResponse = xmlHttp.responseText
End Function

Sub PopulateSheet(json As Object)
    ' Create a new worksheet
    Dim ws As Worksheet
    Set ws = Worksheets.Add

    ' Output headers
    ws.Cells(1, 1).Value = "ID"
    ws.Cells(1, 2).Value = "Name"
    ws.Cells(1, 3).Value = "Username"
    ws.Cells(1, 4).Value = "Email"
    ws.Cells(1, 5).Value = "Phone"

    ' Output data from JSON to the worksheet
    Dim user As Object
    Dim row As Integer
    row = 2 ' Start from row 2 to leave room for headers

    For Each user In json
        ws.Cells(row, 1).Value = user("id")
        ws.Cells(row, 2).Value = user("name")
        ws.Cells(row, 3).Value = user("username")
        ws.Cells(row, 4).Value = user("email")
        ws.Cells(row, 5).Value = user("phone")
        row = row + 1
    Next user
End Sub
