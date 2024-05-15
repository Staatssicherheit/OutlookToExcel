Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As Long, _
    ByVal hWndChildAfter As Long, _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE

Sub CopyDataFromBCGPGridCtrl()
    Dim hwndParent As Long
    Dim hwndGridCtrl As Long
    Dim strBuffer As String
    Dim lngBufferSize As Long

    ' Find the handle of the main application window
    hwndParent = FindWindow(vbNullString, "Your Main Window Title")

    If hwndParent <> 0 Then
        ' Find the handle of the BCGPGridCtrl window
        hwndGridCtrl = FindWindowEx(hwndParent, 0&, "BCGPGridCtrl", vbNullString)

        If hwndGridCtrl <> 0 Then
            ' Get the length of the text in the grid control
            lngBufferSize = SendMessage(hwndGridCtrl, WM_GETTEXTLENGTH, 0&, ByVal 0&) + 1
            strBuffer = Space$(lngBufferSize)

            ' Get the text from the grid control
            SendMessage hwndGridCtrl, WM_GETTEXT, ByVal lngBufferSize, ByVal strBuffer

            ' Now strBuffer contains the text from the grid control
            ' You can process it as needed, for example, copy it to the clipboard or an Excel sheet
        Else
            MsgBox "BCGPGridCtrl window not found."
        End If
    Else
        MsgBox "Main application window not found."
    End If
End Sub
