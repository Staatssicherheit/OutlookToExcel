Sub CopySheets()
    Dim sourceWB As Workbook
    Dim destWB As Workbook
    Dim sourceSheet As Worksheet
    Dim destSheet As Worksheet
    Dim i As Integer
    Dim sheetNames() As String
    Dim copyRanges() As String
    
    ' List of sheet names to copy
    sheetNames = Array("Sheet1", "Sheet2", "Sheet3", "Sheet4", "Sheet5", "Sheet6", "Sheet7")
    
    ' List of ranges to copy for specific sheets
    copyRanges = Array("A1:B10", "C3:E15", "F2:H20", "", "", "", "") ' Adjust the ranges as needed
    
    ' Open both workbooks
    Set sourceWB = Workbooks.Open("Path_to_excel1.xlsx")
    Set destWB = Workbooks.Open("Path_to_excel2.xlsx")
    
    ' Copy sheets from source to destination
    For i = LBound(sheetNames) To UBound(sheetNames)
        Set sourceSheet = sourceWB.Sheets(sheetNames(i))
        If i <= UBound(copyRanges) Then
            ' Copy specified range for the sheet
            If copyRanges(i) <> "" Then
                sourceSheet.Range(copyRanges(i)).Copy
                destWB.Sheets(destWB.Sheets.Count).Range("A1").PasteSpecial xlPasteValues
            Else
                ' If the range is empty, copy the entire sheet
                sourceSheet.Copy Before:=destWB.Sheets(destWB.Sheets.Count)
            End If
        Else
            ' Copy the entire sheet for sheets beyond the specified range copies
            sourceSheet.Copy Before:=destWB.Sheets(destWB.Sheets.Count)
        End If
    Next i
    
    ' Close the source workbook without saving changes
    sourceWB.Close SaveChanges:=False
    
    ' Save and close the destination workbook
    destWB.Save
    destWB.Close SaveChanges:=True
    
    ' Release objects from memory
    Set sourceWB = Nothing
    Set destWB = Nothing
    Set sourceSheet = Nothing
    Set destSheet = Nothing
    
    MsgBox "Sheets copied successfully!"
End Sub
