Sub CopyRangesToWord()

    Dim wordApp As Object
    Dim wordDoc As Object
    Dim ws As Worksheet
    Dim rng As Range
    Dim wordRange As Object
    Dim copyRanges() As Variant
    Dim i As Integer

    ' Define the ranges you want to copy
    copyRanges = Array("Sheet1!A1:B10", "Sheet2!C1:D10", "Sheet3!E1:F10") ' Update with your desired ranges

    ' Create a new instance of Word
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = True ' Make Word visible

    ' Create a new Word document
    Set wordDoc = wordApp.Documents.Add

    ' Loop through each range in the copyRanges array
    For i = LBound(copyRanges) To UBound(copyRanges)
        ' Split the range into sheet name and cell range
        Dim rangeParts() As String
        rangeParts = Split(copyRanges(i), "!")
        Set ws = ThisWorkbook.Sheets(rangeParts(0))
        Set rng = ws.Range(rangeParts(1))

        ' Copy the range
        rng.Copy

        ' Paste the range into the Word document
        Set wordRange = wordDoc.Range
        wordRange.Collapse Direction:=0
        wordRange.Paste

        ' Insert a page break after each pasted range
        wordDoc.Content.InsertAfter vbCrLf & Chr(12)
    Next i

End Sub
