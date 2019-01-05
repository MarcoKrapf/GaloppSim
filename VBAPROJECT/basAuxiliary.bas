Attribute VB_Name = "basAuxiliary"
Option Explicit
Option Private Module

'This module contains auxiliary procedures and general functions


'Place the UserForm in the center of the Window
Public Sub PlaceUserFormInCenter(frmMe As Object)
        '[ frmMe as UserForm does not work:
        '  http://www.office-loesung.de/ftopic516038_0_0_asc.php ]
    With frmMe
        .StartUpPosition = 0
        .top = ActiveWindow.top + ((ActiveWindow.Height - frmMe.Height) / 2)
        .left = ActiveWindow.left + ((ActiveWindow.Width - frmMe.Width) / 2)
    End With
End Sub

'Freezing/unfreezing the window pane
Public Sub Freeze(col As Integer, row As Integer, direction As Boolean)
    With Application.ActiveWindow
        .SplitColumn = col
        .SplitRow = row
        .FreezePanes = direction
    End With
End Sub

'Scrolling
Public Sub Scroll(col As Integer, row As Integer)
    On Error Resume Next
        With Application.ActiveWindow
            .ScrollColumn = col
            .ScrollRow = row
        End With
    On Error GoTo 0
End Sub

'Return the column number
Public Function GetColumn(wks As Worksheet, search As String) As Integer
    Dim c As Integer
    For c = 1 To 16384 'Number of columns on a worksheet
        If UCase(wks.Cells(1, c).Value) = UCase(search) Then 'Look in row 1
            GetColumn = c
            Exit Function
        End If
    Next c
End Function

'Return the row number
Public Function GetRow(wks As Worksheet, search As String) As Long
    Dim r As Long
    For r = 1 To 1048576 'Number of columns on a worksheet
        If UCase(wks.Cells(r, 1).Value) = UCase(search) Then 'Look in column 1
            GetRow = r
            Exit Function
        End If
    Next r
End Function

'Returns a text component according to the language
Public Function GetText(arr As Variant, id As String) As String
    Dim r As Long
    For r = 1 To UBound(arr, 2)
        If arr(0, r) = id Then
            GetText = arr(1, r)
            Exit Function
        End If
    Next r
End Function

'Pop-up showing an error message
Public Sub CodeCrash()
    'Set the button mode
    g_strMsgButtons = "OK"
    'Assign the text for the pop-up
    g_strMsgCaption = g_c_tool
    g_strMsgText = GetText(g_arr_Text, "ERROR001")
    'Display the pop-up (modal)
    frmMsg_MultiPurpose.show (vbModal)
End Sub

'Colour picker
Public Function ColPick(current As Long) As Long
    Dim returncode As Integer
    returncode = Application.Dialogs(xlDialogEditColor).show(10)
    If returncode <> 0 Then
        ColPick = ActiveWorkbook.Colors(10)
    Else
        ColPick = current
    End If
End Function

'Formatting: Race information on the worksheet
Public Sub RaceInfoWorksheet(colBack As Long, colFore As Long, top, show As Boolean)
    'Colours
    With g_wksRace.Range(Cells(1 + top, 1), Cells(3 + top, 12))
        .Interior.color = colBack
        .Font.color = colFore
    End With
    'Leader
    With g_wksRace.Cells(1 + top, 2)
        .Font.name = "Arial Black"
        .Font.Size = 8
        .Font.Bold = True
        .Value = ""
    End With
    With g_wksRace.Cells(2 + top, 10)
        .Font.name = "Arial Black"
        .Font.Size = 11
        .Font.Bold = True
        .Value = ""
    End With
    'Race distance (metres run)
    With g_wksRace.Cells(3 + top, 11)
        .Font.name = "Arial Black"
        .Font.Size = objOption.ZOOM_LEVEL + 5
        .IndentLevel = 1 'Text indented
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Value = ""
    End With
    'Race progress bar
    If objOption.RACE_INFO_PROGRESS Then
        With g_wksRace.Cells(3 + top, 12)
            .Font.name = "Arial"
            .Font.Size = 11
            .Value = ""
            If show = True Then
                .Interior.color = colFore
                .Font.color = colBack
                .BorderAround Weight:=xlThick, color:=colFore 'Draw cell frame
            Else
                .Borders.LineStyle = xlNone 'Delete cell frame
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

'Pop-up to delock algorithms
Public Function AllowAlgorithms() As Boolean
    'Set the button mode
    g_strMsgButtons = "YesNo"
    'Assign the text for the pop-up
    g_strMsgCaption = GetText(g_arr_Text, "USERFORM007")
    g_strMsgText = GetText(g_arr_Text, "ERROR007")
    'Display the pop-up
    frmMsg_MultiPurpose.show (vbModal) 'modal
    'Evaluate the return value
    If g_strButtonPressed = "YES" Then
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

'Converting a colour code from Excel long format to a grey tone in RGB format
Public Function RGBtoGrey(lngColor As Long) As Long
    RGBtoGrey = ( _
                (lngColor Mod 256) + _
                ((lngColor \ 256) Mod 256) + _
                ((lngColor \ 256 ^ 2) Mod 256) _
                ) / 3
End Function

'Converting a grey tone in RGB format to the corresponding grey code
'in Excel long format
Public Function GreyToLong(lngGrey As Long) As Long
    GreyToLong = lngGrey + lngGrey * 256 + lngGrey * 256 ^ 2
End Function

'Speech output
Public Sub SpeechOut(words As String)
    Application.SPEECH.Speak (words), SpeakAsync:=True
End Sub
