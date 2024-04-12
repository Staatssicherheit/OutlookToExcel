
Sub CopyToExcel()
Dim xlApp As Object
Dim xlWB As Object
Dim xlSheet As Object
Dim olItem As Outlook.MailItem
Dim vText As Variant
Dim sText As String
Dim vItem As Variant
Dim vAddr As Variant
Dim oRng As Object
Dim i As Long, j As Long
Dim rCount As Long
Dim sAddr As String
Dim bXstarted As Boolean
Const strPath As String = "D:\My Documents\Workbook Name.xlsx"        'the path of the workbook

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "No Items selected!", vbCritical, "Error"
        Exit Sub
    End If
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Err <> 0 Then
        Application.StatusBar = "Please wait while Excel source is opened ... "
        Set xlApp = CreateObject("Excel.Application")
        bXstarted = True
    End If
    On Error GoTo 0
    'Open the workbook to input the data
    Set xlWB = xlApp.Workbooks.Open(strPath)
    Set xlSheet = xlWB.Sheets("Sheet1")

    'Process each selected record
    For Each olItem In Application.ActiveExplorer.Selection
        sText = olItem.Body
        vText = Split(sText, Chr(13))
        'Find the next empty line of the worksheet
        rCount = xlSheet.Range("A" & xlSheet.Rows.Count).End(-4162).Row
        rCount = rCount + 1
       
        'Check each line of text in the message body
        For i = UBound(vText) To 0 Step -1
            If InStr(1, vText(i), "Student name:") > 0 Then
                vItem = Split(vText(i), Chr(58))
                xlSheet.Range("A" & rCount) = Trim(vItem(1))
            End If

            If InStr(1, vText(i), "Project name:") > 0 Then
                vItem = Split(vText(i), Chr(58))
                xlSheet.Range("B" & rCount) = Trim(vItem(1))
            End If

            If InStr(1, vText(i), "Booth Code:") > 0 Then
                vItem = Split(vText(i), Chr(58))
                xlSheet.Range("C" & rCount) = Trim(vItem(1))
            End If

            If InStr(1, vText(i), "Date completed:") > 0 Then
                vItem = Split(vText(i), Chr(58))
                xlSheet.Range("D" & rCount) = Trim(vItem(1))
            End If

            If InStr(1, vText(i), "Time completed:") > 0 Then
                vItem = Split(vText(i), Chr(58))
                xlSheet.Range("E" & rCount) = Trim(vItem(1))
            End If

            If InStr(1, vText(i), "Total time spent in course:") > 0 Then
                vItem = Split(vText(i), Chr(58))
                xlSheet.Range("F" & rCount) = Trim(vItem(1))
            End If
        Next i
        xlWB.Save
    Next olItem
    xlWB.Close SaveChanges:=True
    If bXstarted Then
        xlApp.Quit
    End If
    Set xlApp = Nothing
    Set xlWB = Nothing
    Set xlSheet = Nothing
    Set olItem = Nothing
End Sub
