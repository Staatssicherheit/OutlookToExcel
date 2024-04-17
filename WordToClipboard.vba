Sub CopyAreaWithTable()
    Dim doc As Document
    Dim rng As Range
    
    ' Open the Word document
    Set doc = Documents.Open("X:\testingparagraph.docx")
    
    ' Define the range of the area you want to copy
    Set rng = doc.Range(Start:=doc.Paragraphs(1).Range.Start, End:=doc.Paragraphs(3).Range.End)
    
    ' Copy the range
    rng.Copy
    
    ' Close the document without saving changes
    doc.Close SaveChanges:=wdDoNotSaveChanges
End Sub
