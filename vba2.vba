Sub OutlookToExcel()
Dim oLookInspector As Outlook.Inspector
Dim oLookMailitem As Outlook.MailItem
Dim oLookWordDoc As Word.Document
Dim oLookWordTbl As Word.Table
Dim xlApp As Application
Dim xlBook As Workbook
Dim xlWrkSheet As Worksheet
Dim OutlookApp As Outlook.Application
Dim OutlookNamespace As Namespace
Dim Folder As MAPIFolder
Set OutlookApp = New Outlook.Application
Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
Set Folder = OutlookNamespace.GetDefaultFolder(olFolderInbox).Folders("TataMutualFund")
'Set oLookMailitem = Outlook.Application.ActiveExplorer.CurrentFolder.Items("data")
Set oLookMailitem = Folder.Items(3)
Set oLookInspector = oLookMailitem.GetInspector
Debug.Print oLookInspector.CurrentItem
Set oLookWordDoc = oLookInspector.WordEditor
Set xlApp = New Excel.Application
    xlApp.Visible = True

Set xlBook = xlApp.Workbooks.Add
Set xlWrkSheet = xlBook.Worksheets.Add
Set oLookWordTbl = oLookWordDoc.Tables(1)
    oLookWordTbl.Range.Copy
    xlWrkSheet.Paste Destination:=xlWrkSheet.Range("A1")

End Sub
