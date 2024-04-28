Sub CopySheets()
    Dim sourceWB As Workbook
    Dim destWB As Workbook
    Dim sourceSheet As Worksheet
    Dim destSheet As Worksheet
    Dim i As Integer
    
    ' Open both workbooks
    Set sourceWB = Workbooks.Open("Path_to_excel1.xlsx")
    Set destWB = Workbooks.Open("Path_to_excel2.xlsx")
    
    ' Copy sheets from source to destination
    For Each sourceSheet In sourceWB.Sheets
        ' Copy each sheet to the destination workbook before the last sheet
        sourceSheet.Copy Before:=destWB.Sheets(destWB.Sheets.Count)
    Next sourceSheet
    
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
