Sub CopyToClipboardWithSentences()
    Dim myData As String
    Dim ws As Worksheet
    Dim rng As Range
    Dim i As Integer
    
    ' Array of sheet names and corresponding ranges
    Dim sheetsAndRanges As Variant
    sheetsAndRanges = Array("메인_보고서!B5:R39", "고객사 단말현황!D3:X24", "고객사점검!B2:M49", "대표번호 녹취내역")
    
    ' Copy each specified range to clipboard
    For i = LBound(sheetsAndRanges) To UBound(sheetsAndRanges)
        If InStr(sheetsAndRanges(i), "!") > 0 Then
            Set ws = ThisWorkbook.Sheets(Split(sheetsAndRanges(i), "!")(0))
            Set rng = ws.Range(Split(sheetsAndRanges(i), "!")(1))
        Else
            Set ws = ThisWorkbook.Sheets(sheetsAndRanges(i))
            Set rng = ws.UsedRange
        End If
        rng.Copy
        ' Add corresponding sentence after the 2nd, 3rd, and 4th arrays
        If i = 1 Then
            myData = myData & vbCrLf & "6. 고객사 단말 현황"
        ElseIf i = 2 Then
            myData = myData & vbCrLf & "7. 고객사별 서비스점검 내역"
        ElseIf i = 3 Then
            myData = myData & vbCrLf & "8. 대표번호 녹취내역"
        End If
        myData = myData & vbCrLf & sheetsAndRanges(i)
    Next i
    
    ' Copy the data with sentences to clipboard
    With New MSForms.DataObject
        .SetText myData
        .PutInClipboard
    End With
End Sub
