Attribute VB_Name = "basAuxiliary"
Option Explicit
Option Private Module

'This module contains public accessible procedures and general functions
'which can be called by procedures in other modules
'   Module basAuxiliary

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Get the screen width of the main screen
Public Function GetScreenWidth() As Variant
    GetScreenWidth = GetSystemMetrics(0) 'Width (pixels)
End Function
'Get the screen height of the main screen
Public Function GetScreenHeight() As Variant
    GetScreenHeight = GetSystemMetrics(1) 'Height (pixels)
End Function

'Open a website in the standard browser (if possible)
Public Sub OpenURL(URL As String)
    On Error Resume Next 'Ignore the command if an error occures
    ActiveWorkbook.FollowHyperlink Address:=URL
    On Error GoTo 0 'Disable the error handler (not necessary as this is the end of the procedure)
End Sub
    
'Send an e-mail (if possible)
Public Sub SendMail()
    Dim objMail As Object 'Shell object for the e-mail
    On Error Resume Next
    Set objMail = CreateObject("Shell.Application")
    objMail.ShellExecute "mailto:" & g_c_email _
        & "&subject=" & g_c_tool
End Sub

'Place the pop-up in the center of the window
Public Sub PlaceUserFormInCenter(frmMe As Object)
        '[ frmMe as UserForm does not work:
        '  http://www.office-loesung.de/ftopic516038_0_0_asc.php ]
    With frmMe
        .StartUpPosition = 0
        .top = ActiveWindow.top + ((ActiveWindow.Height - frmMe.Height) / 2) 'Align vertically
        .left = ActiveWindow.left + ((ActiveWindow.width - frmMe.width) / 2) 'Align horizontally
    End With
End Sub

'Freeze/unfreeze the window pane
Public Sub Freeze(col As Integer, row As Integer, direction As Boolean)
    With Application.ActiveWindow
        .SplitColumn = col
        .SplitRow = row
        .FreezePanes = direction 'True = freeze False = unfreeze
    End With
End Sub

'Scrolling
Public Sub Scroll(col As Integer, row As Integer)
    On Error Resume Next
        With Application.ActiveWindow
            .ScrollColumn = col
            .ScrollRow = row
        End With
End Sub

'Force drawing on the GALOPPSIM worksheet
Public Sub ActivateRaceSheet()
    'Two possible "If" statement variants:
    'In a single line (for short, simple tests with only one command if the condition is true)
    If ActiveSheet.name <> g_wksRace.name Then g_wksRace.Activate
    
'    '...or by using the block form syntax with multiple lines that ends with "End If"
'    'for more flexibility like executing several commands if the condition is true
'    If ActiveSheet.name <> g_wksRace.name Then
'        g_wksRace.Activate
'    End If
End Sub

'Return the column number of a search string to be found in row 1
Public Function GetColumn(wks As Worksheet, search As String) As Integer
    Dim c As Integer
    For c = 1 To 16384 'Maximum number of columns on a worksheet as of Excel 2007
        If UCase(wks.Cells(1, c).Value) = UCase(search) Then 'Search in the top row (1)
            GetColumn = c 'Return the column number in which the search string is found
            Exit Function
        End If
    Next c
End Function

'Return the row number of a search string to be found in column A
Public Function GetRow(wks As Worksheet, search As String) As Long
    Dim r As Long
    For r = 1 To 1048576 'Maximum number of rows on a worksheet as of Excel 2007
        If UCase(wks.Cells(r, 1).Value) = UCase(search) Then 'Search in the left column (A)
            GetRow = r 'Return the row number in which the search string is found
            Exit Function
        End If
    Next r
End Function

'Get all text components according to the selected language
Public Sub GetTextComponents()
    Dim i As Integer, col As Integer
    
    'Read the text components into a 2-dimensional Array
        ReDim g_arr_Text(0 To 1, 1 To 2000) 'Resize the Array: 2 columns with 2000 rows
        col = GetColumn(g_wksTEXT, objOption.language) 'get the column with the language
        For i = 1 To UBound(g_arr_Text, 2)
            g_arr_Text(0, i) = g_wksTEXT.Cells(i, 1).Value 'ID taken from column A
            g_arr_Text(1, i) = g_wksTEXT.Cells(i, col).Value 'Text according to the selected language
        Next i
    
End Sub

'Return a single text component from the Array with all text components
Public Function GetText(arr As Variant, id As String) As String
    Dim r As Long
    For r = 1 To UBound(arr, 2)
        If arr(0, r) = id Then
            GetText = arr(1, r)
            Exit Function
        End If
    Next r
End Function

Public Sub PaintPicture(wksSource As Worksheet, wksTarget As Worksheet, ByVal pic As String, cols As Integer, rows As Integer, top As Integer, left As Integer)
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Dim i As Long 'The 'local' variable i on procedure-level overrides the i from the calling procedure so there will be no trouble
    Dim j As Long, k As Long, m As Long 'If you write 'Dim j, k, m As Long' only m will be of type Long, all others are of type Variant
    
    Application.ScreenUpdating = False 'Deactivate screen updating
        
    'Paint a picture
        k = GetColumn(wksSource, pic) 'Get the column with the colour data from the Worksheet "PIC"
        m = 2 'Initial row for reading the colour data for the picture
        
        For i = top To rows
            For j = left To cols
                With wksTarget.Cells(i, j)
                    .Clear 'Clear the cell content
                    .Interior.color = wksSource.Cells(m, k).Value 'Format the cell with the background colour
                End With
                m = m + 1 'Next row on the Worksheet "PIC"
            Next j
        Next i

    Application.ScreenUpdating = True 'Activate screen updating
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "PaintPicture()")
    Call basAuxiliary.CodeCrash
End Sub

'Pop-up showing an error message in case of a runtime error
Public Sub CodeCrash()
    Call ShowMessagePopup(g_c_tool, GetText(g_arr_Text, "ERROR001"), _
        enumButton.OK, vbModal)
End Sub

'Colour picker
Public Function ColPick(current As Long) As Long
    Dim returncode As Integer
    returncode = Application.Dialogs(xlDialogEditColor).show(10)
    If returncode <> 0 Then 'If a colour has been picked
        ColPick = ActiveWorkbook.Colors(10)
    Else 'Click on the "Cancel" button
        ColPick = current
    End If
End Function

'Formattings: Race information on the worksheet
Public Sub RaceInfoWorksheet(colBack As Long, colFore As Long, top, show As Boolean)
    'Colours
    With g_wksRace.Range(Cells(1 + top, 1), Cells(3 + top, 12))
        .Interior.color = colBack
        .Font.color = colFore
    End With
    'Current leader in the race
    With g_wksRace.Cells(1 + top, 2)
        .Font.name = "Arial Black"
        .Font.size = 8
        .Font.Bold = True
        .Value = ""
    End With
    With g_wksRace.Cells(2 + top, 10)
        .Font.name = "Arial Black"
        .Font.size = 11
        .Font.Bold = True
        .Value = ""
    End With
    'Race distance (metres run)
    With g_wksRace.Cells(3 + top, 11)
        .Font.name = "Arial Black"
        .Font.size = objOption.ZOOM_LEVEL + 5
        .IndentLevel = 1 'Text indented
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Value = ""
    End With
    'Race progress bar
    If objOption.RACE_INFO_PROGRESS Then
        With g_wksRace.Cells(3 + top, 12)
            .Font.name = "Arial"
            .Font.size = 11
            .Value = ""
            If show = True Then
                .Interior.color = colFore
                .Font.color = colBack
                .BorderAround Weight:=xlThick, color:=colFore 'Draw a cell frame
            Else
                .Borders.LineStyle = xlNone 'Delete the cell frame
            End If
        End With
    End If
End Sub

'Determine the caption of the start button
Public Function getCaptionStartBtn(bettingMode As Boolean) As String
    If bettingMode = False Then
        getCaptionStartBtn = "BTN003a" '"Start the race"
    Else
        getCaptionStartBtn = "BTN003b" '"Betting and race"
    End If
End Function

'Pop-up for re-activation of all algorithms
Public Function AllowAlgorithms() As Boolean
    Call ShowMessagePopup(GetText(g_arr_Text, "USERFORM007"), _
        GetText(g_arr_Text, "ERROR007"), _
        enumButton.YesNo, vbModal)
    
    'Evaluate the return value
    If g_enumButton = enumButton.yes Then
        AllowAlgorithms = False
    Else
        AllowAlgorithms = True
    End If
End Function

'Place the cursor far away (in the upper right corner of the screen)
Public Sub CursorAway()
    On Error Resume Next
        g_wksRace.Cells(1, ActiveWindow.VisibleRange.Columns.count - 1).Activate
End Sub

'Conversion of a colour code from Excel data type 'Long' to a grey tone in RGB format
Public Function RGBtoGrey(lngColor As Long) As Long
    RGBtoGrey = ( _
                (lngColor Mod 256) + _
                ((lngColor \ 256) Mod 256) + _
                ((lngColor \ 256 ^ 2) Mod 256) _
                ) / 3
End Function

'Conversion of a grey tone in RGB format to the corresponding grey code
'in Excel data type 'Long'
Public Function GreyToLong(lngGrey As Long) As Long
    GreyToLong = lngGrey + lngGrey * 256 + lngGrey * 256 ^ 2
End Function

'Voice output of a text
Public Sub SpeechOut(words As String)
    Application.SPEECH.Speak (words), SpeakAsync:=True
End Sub

'Simple custom pop-up for information and attention messages
Public Sub ShowInfoPopup(caption As String, text As String, _
            attention As Boolean, mode As Long, _
            Optional fontsize As Integer)
            
    With frmMsg_Info
        'Popup caption
        .caption = caption
        
        'Show one out of two icons for either an information or a warning message
        With .imgAttention
        .Visible = attention 'Black exclamation mark in a yellow triangle
        .left = 6
        End With
        With .imgInformation 'White exclamation mark in a blue circle
        .Visible = Not attention
        .left = 6
        End With
        
        'Text label
        With .lblText
        .top = 12
        .left = 78
        .width = 800 'Initial width
        .Height = 800 'Initial height
        If fontsize > 0 Then .Font.size = fontsize 'If not submitted: 8.25 (standard value)
        .caption = text  'Text content
        .AutoSize = True 'Shrink the label to the optimal size
        End With
        
        'Adjust the size of the pop-up and show it
        .width = .lblText.width + 100
        .Height = .lblText.Height + 100
        .show (mode) '1 = vbModal // 0 = vbModeless
    End With
End Sub

'Pop-up for a custom input box
Public Sub ShowInputPopup(caption As String, text As String, boxWidth As Integer, _
    maxLength, buttons As Long, mode As Long)
    With frmInp_MultiPurpose
        'Popup caption
        .caption = caption
        
        'Set all buttons invisible
        .cmdInpOK.Visible = False
        .cmdInpCancel.Visible = False
        
        'Text label that describes the expected input
        With .lblInp01
        .top = 12
        .left = 12
        .caption = text
        .AutoSize = True
        End With
        
        'Box for user input
        With .txtInp01
            .left = frmInp_MultiPurpose.lblInp01.width + 18 'Alignment dependent on the decription label
            .width = boxWidth
            .Height = 20
            .maxLength = maxLength 'Maximum number of characters allowed
        End With
        
        'Buttons
        Select Case buttons
            Case enumButton.OK 'Only OK button
                Call AlignButtonInp(.cmdInpOK, GetText(g_arr_Text, "BTN014"), 0)
            Case enumButton.CancelOK 'OK and Cancel button
                Call AlignButtonInp(.cmdInpOK, GetText(g_arr_Text, "BTN014"), 0)
                Call AlignButtonInp(.cmdInpCancel, GetText(g_arr_Text, "BTN015"), .cmdInpOK.width + 5)
        End Select
        
        'Adjust the size of the pop-up and show it
        .width = .txtInp01.left + .txtInp01.width + 20
        .Height = .lblInp01.Height + 92
        .show (mode) '1 = vbModal // 0 = vbModeless
    End With
End Sub

'Alignment of a button in a custom input box
Private Sub AlignButtonInp(cmdButton As Object, strText As String, intCorrection As Integer)
    With cmdButton
        .Visible = True
        .caption = strText
        .top = frmInp_MultiPurpose.txtInp01.top + frmInp_MultiPurpose.txtInp01.Height + 12
        .left = frmInp_MultiPurpose.txtInp01.left + frmInp_MultiPurpose.txtInp01.width - .width - intCorrection
    End With
End Sub

'Pop-up for an advanced custom message box with buttons
Public Sub ShowMessagePopup(caption As String, text As String, buttons As Long, mode As Long)
    With frmMsg_MultiPurpose
        'Popup caption
        .caption = caption

        'Set all buttons invisible
        .cmdMsgOK.Visible = False
        .cmdMsgCancel.Visible = False
        .cmdMsgYes.Visible = False
        .cmdMsgNo.Visible = False
        
        'Background label
        With .lblMsg02
            .BackColor = &HFFFFFF 'White
            .caption = ""
            .top = 0
            .left = 0
        End With
        
        'Text label
        With .lblMsg01
            .BackColor = &HFFFFFF 'White
            .top = 12
            .left = 12
            .caption = text
            .AutoSize = True
        End With
        
        'Adjust the size of the background label
        With .lblMsg02
            .Height = frmMsg_MultiPurpose.lblMsg01.Height + 30
            .width = frmMsg_MultiPurpose.lblMsg01.width + 35
        End With
        
        'Adjust the buttons
        Select Case buttons
            Case enumButton.OK
                Call AlignButtonMsg(.cmdMsgOK, GetText(g_arr_Text, "BTN014"), 0)
            Case enumButton.CancelOK
                Call AlignButtonMsg(.cmdMsgOK, GetText(g_arr_Text, "BTN014"), 0)
                Call AlignButtonMsg(.cmdMsgCancel, GetText(g_arr_Text, "BTN015"), .cmdMsgOK.width + 5)
            Case enumButton.YesNo
                Call AlignButtonMsg(.cmdMsgYes, GetText(g_arr_Text, "BTN016"), .cmdMsgNo.width + 5)
                Call AlignButtonMsg(.cmdMsgNo, GetText(g_arr_Text, "BTN017"), 0)
        End Select
        
        'Adjust the size of the pop-up and show it
        .width = .lblMsg01.width + 35
        .Height = .lblMsg01.Height + 105
        .show (mode) '1 = vbModal // 0 = vbModeless
        
    End With
End Sub

'Alignment of a button in an advanced custom message box
Private Sub AlignButtonMsg(cmdButton As Object, strText As String, intCorrection As Integer)
    With cmdButton
        .Visible = True
        .caption = strText
        .top = frmMsg_MultiPurpose.lblMsg01.top + frmMsg_MultiPurpose.lblMsg01.Height + 30
        .left = frmMsg_MultiPurpose.lblMsg01.left + frmMsg_MultiPurpose.lblMsg01.width - .width - intCorrection
    End With
End Sub
