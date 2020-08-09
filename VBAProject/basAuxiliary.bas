Attribute VB_Name = "basAuxiliary"
Option Explicit
Option Private Module

'This module contains public accessible procedures and general functions
'which can be called by procedures in other modules
'   Module basAuxiliary

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
    objMail.ShellExecute "mailto:" & g_c_email
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

'Apply either auto fit or a selected zoom level
Public Sub AutoZoom(usecase As String, Optional percentage As Integer)

    Application.ScreenUpdating = False 'Deactivate screen updating
    ActiveWindow.zoom = 100 'Reset in case the zoom is already too small
    
    Select Case usecase
        Case "FIX" 'Zoom to the desired percentage value
            ActiveWindow.zoom = percentage
        Case "Race" 'Auto fit: Reduce the zoom level until the advertisements below the race track fit in the screen
            Do While ActiveWindow.VisibleRange.rows.count < rows(objRace.NUMBER_ENROLLED * 2 + 19 + objBasicData.TOP_ROWS).row
                If ActiveWindow.zoom > 10 Then ActiveWindow.zoom = ActiveWindow.zoom - 1
            Loop
        Case "FinishPhoto" 'Auto fit: Reduce the zoom level until the entire finish photo is displayed
            Do While ActiveWindow.VisibleRange.rows(ActiveWindow.VisibleRange.rows.count).row <= (objRace.NUMBER_ENROLLED * 2 + 8 + objBasicData.TOP_ROWS) _
                Or ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column <= (objRace.METRES + objBasicData.LEFT_COLS + 175 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR))
                
                If ActiveWindow.zoom > 10 Then ActiveWindow.zoom = ActiveWindow.zoom - 1
            Loop
        Case "RankingList" 'Auto fit: Reduce the zoom level until the entire ranking list is displayed
            Do While ActiveWindow.VisibleRange.rows(ActiveWindow.VisibleRange.rows.count).row <= objRace.NUMBER_ENROLLED * 2 + 20 + objRace.NUMBER_STARTING + 1 + objBasicData.TOP_ROWS _
                Or ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column <= objRace.METRES + objBasicData.LEFT_COLS + 175 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)
                
                If ActiveWindow.zoom > 10 Then ActiveWindow.zoom = ActiveWindow.zoom - 1
            Loop
        Case "WinnerPhoto" 'Auto fit: Reduce the zoom level until all winner photos are displayed
            Do While ActiveWindow.VisibleRange.rows(ActiveWindow.VisibleRange.rows.count).row <= objRace.NUMBER_ENROLLED * 2 + 40 + objBasicData.TOP_ROWS _
                Or ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column <= objRace.METRES + objBasicData.LEFT_COLS + 196 + (objStat.WINNERS * 21) + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR)

                    If ActiveWindow.zoom > 10 Then ActiveWindow.zoom = ActiveWindow.zoom - 1
                    Call Scroll(objRace.METRES + objBasicData.LEFT_COLS + 16 + 160 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR), objBasicData.TOP_ROWS + (objRace.NUMBER_ENROLLED * 2 + 9))
            Loop
        Case "Movie" 'Auto fit: Reduce the zoom level until the complete movie is displayed
            Do While ActiveWindow.VisibleRange.rows.count < 40 Or ActiveWindow.VisibleRange.Columns.count < 100
                If ActiveWindow.zoom > 10 Then ActiveWindow.zoom = ActiveWindow.zoom - 1
            Loop
        Case Else 'If none of the optional argumets is submitted
            MsgBox "No zoom performed." & vbNewLine _
                    & "Check the arguments for calling this function."
    End Select
    
    Application.ScreenUpdating = True 'Activate screen updating

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
    
    'Variables used for the ecstasy colour mode
    Dim colHeavenPop As Long, colGrassPop As Long, colRandomPop As Long
    colHeavenPop = PopArtColour(Int((16777215 - 0 + 1) * Rnd + 0))
    colGrassPop = PopArtColour(Int((16777215 - 0 + 1) * Rnd + 0))
    
    'Variables used for the Random colour mode
    Dim colHeaven As Long, colGrass As Long, colRandom As Long
    colHeaven = Int((16777215 - 0 + 1) * Rnd + 0)
    colGrass = Int((16777215 - 0 + 1) * Rnd + 0)
            
    Application.ScreenUpdating = False 'Deactivate screen updating
        
    'Paint a picture
        k = GetColumn(wksSource, pic) 'Get the column with the colour data from the Worksheet "PIC"
        m = 2 'Initial row for reading the colour data for the picture
        
        For i = top To top + rows - 1
            For j = left To left + cols - 1
                With wksTarget.Cells(i, j)
                    .Clear 'Clear the cell content
                    .Interior.color = wksSource.Cells(m, k).Value 'Format the cell with the background colour
                
                    'Take the colour mode into account
                    Select Case g_strColourMode
                        Case "POPART"
                            If wksTarget.Cells(i, j).Interior.color > 0 And wksTarget.Cells(i, j).Interior.color < 16777215 Then _
                            wksTarget.Cells(i, j).Interior.color = PopArtColour(wksTarget.Cells(i, j).Interior.color)
                        Case "LSD"
                            With wksTarget.Cells(i, j)
                                Select Case .Interior.color
                                    Case 0, 16777215 'No change
                                    
                                    Case 14726300 'Heaven
                                        .Interior.color = colHeavenPop
                                    Case 9359529 'Grass
                                        .Interior.color = colGrassPop
                                    Case Else
                                        Do
                                            colRandomPop = PopArtColour(Int((16777215 - 0 + 1) * Rnd + 0))
                                        Loop Until colRandomPop <> colHeavenPop And colRandomPop <> colGrassPop
                                        .Interior.color = colRandomPop
                                End Select
                            End With
                        Case "SMARTIES"
                            With wksTarget.Cells(i, j)
                                Select Case .Interior.color
                                    Case 0, 16777215 'No change
                                    
                                    Case 14726300 'Heaven
                                        .Interior.color = colHeaven
                                    Case 9359529 'Grass
                                        .Interior.color = colGrass
                                    Case Else
                                        Do
                                            colRandom = Int((16777215 - 0 + 1) * Rnd + 0)
                                        Loop Until colRandom <> colHeaven And colRandom <> colGrass
                                        .Interior.color = colRandom
                                End Select
                            End With
                        Case "TV1960"
                            wksTarget.Cells(i, j).Interior.color = GreyToLong(CInt(RGBtoGrey(CLng(wksTarget.Cells(i, j).Interior.color))))
                        Case "DARKMODE"
                            Select Case .Interior.color
                                Case 0  'Black: change to white
                                    wksTarget.Cells(i, j).Interior.color = 16777215
                                Case 16777215  'White: change to black
                                    wksTarget.Cells(i, j).Interior.color = 0
'                                Case 0, 16777215 'No change
                                
                                Case 14726300 'Heaven
                                    .Interior.color = 2697513 'Dark grey
                                Case 52377 'Grass
                                    .Interior.color = 0 'Black
                                Case Else
                                    wksTarget.Cells(i, j).Interior.color = DarkModeColour(wksTarget.Cells(i, j).Interior.color)
                            End Select
                        Case "24H"
                            With wksTarget.Cells(i, j)
                                Select Case .Interior.color
                                    Case 14726300 'Heaven
                                        .Interior.color = objOption.DAYLIGHT_COL
                                    Case Else
                                        wksTarget.Cells(i, j).Interior.color = DuskDawn(wksTarget.Cells(i, j).Interior.color, Abs(22 * objOption.DAYLIGHT))
                                End Select
                            End With
                    End Select
                
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
Public Sub RaceInfoWorksheet(colBack As Long, colFore As Long, topRows, show As Boolean)
    
    'Colours
    With g_wksRace.Range(Cells(1 + topRows, 1), Cells(3 + topRows, 13))
        .Interior.color = colBack
        .Font.color = colFore
        If objRace.SPECIAL = "PARTICULATES" Then .Interior.Pattern = objOption.PARTICULATES_PATTERN
    End With
    
    'Current leader in the race
    With g_wksRace.Cells(1 + topRows, 2) '"The current leader is"
        .Font.name = "Arial Black"
        .Font.size = 8
        .Font.Bold = True
        .Value = ""
    End With
    With g_wksRace.Cells(2 + topRows, 11) 'Horse name
        .Font.name = "Arial Black"
        .Font.size = 11
        .Font.Bold = True
        .Value = ""
    End With
    
    'Race distance (metres run)
    With g_wksRace.Cells(3 + topRows, 11)
        .Font.name = "Arial Black"
        .Font.size = 8
        .IndentLevel = 1 'Text indented
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Value = ""
    End With
    
    'Race progress bar
    If objOption.RACE_INFO_PROGRESS Then
    
        If show = True Then
    
            Dim t As Double, h As Double, l As Double, w As Double
            With Cells(3 + topRows, 12)
                t = .top
                h = .Height
                l = .left
                w = .width - 1 'Show the full frame when scrolling
            End With
            
            Set g_shpFrame = g_wksRace.Shapes.AddShape(msoShapeRectangle, l, t, w, h)

            With g_shpFrame
                .Line.Weight = 2
                .Line.ForeColor.RGB = RGB( _
                            objOption.RACE_INFO_COL_F Mod 256, _
                            (objOption.RACE_INFO_COL_F \ 256) Mod 256, _
                            (objOption.RACE_INFO_COL_F \ 256 \ 256) Mod 256)
                                'Extract RGB values from Long value
                .Fill.ForeColor.RGB = RGB( _
                            objOption.RACE_INFO_COL_B Mod 256, _
                            (objOption.RACE_INFO_COL_B \ 256) Mod 256, _
                            (objOption.RACE_INFO_COL_B \ 256 \ 256) Mod 256)
                                'Extract RGB values from Long value
                .name = "shpRaceProgressFrame"
            End With
    
            Set g_shpBar = g_wksRace.Shapes.AddShape(msoShapeRectangle, l, t, 0, h)
            With g_shpBar
                .Line.Visible = msoFalse
                .Fill.ForeColor.RGB = RGB( _
                            objOption.RACE_INFO_COL_F Mod 256, _
                            (objOption.RACE_INFO_COL_F \ 256) Mod 256, _
                            (objOption.RACE_INFO_COL_F \ 256 \ 256) Mod 256)
                                'Extract RGB values from Long value
                .name = "shpRaceProgressBar"
            End With
            
        Else
            g_shpBar.Delete
            g_shpFrame.Delete
            
        End If
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

'Extract RED value
Public Function GetRed(lngColor As Long)
    GetRed = lngColor Mod 256
End Function
'Extract GREEN value
Public Function GetGreen(lngColor As Long)
    GetGreen = (lngColor \ 256) Mod 256
End Function
'Extract BLUE value
Public Function GetBlue(lngColor As Long)
    GetBlue = (lngColor \ 256 ^ 2) Mod 256
End Function

Public Function DuskDawn(ByVal lngColor As Long, intPercentage As Integer)
    DuskDawn = RGB(GetRed(lngColor) * (100 - intPercentage) / 100, _
                    GetGreen(lngColor) * (100 - intPercentage) / 100, _
                    GetBlue(lngColor) * (100 - intPercentage) / 100)
End Function

Public Function BrightLight(ByVal lngColor As Long, intPercentage As Integer)
    BrightLight = RGB(GetRed(lngColor) + (255 - GetRed(lngColor)) * intPercentage / 100, _
                    GetGreen(lngColor) + (255 - GetGreen(lngColor)) * intPercentage / 100, _
                    GetBlue(lngColor) + (255 - GetBlue(lngColor)) * intPercentage / 100)
End Function

'Function that returns the nearest PopArt colour
'to a given input colour
Public Function PopArtColour(ByVal inputColour As Long) As Long
    Dim i As Integer
    Dim lngPopArt(0 To 7) As Long 'Array mit den Farben

    lngPopArt(0) = 39423 'Orange (dummy)
    lngPopArt(1) = 39423 'Orange
    lngPopArt(2) = 65280 'Green
    lngPopArt(3) = 195581 'Yellow
    lngPopArt(4) = 6684927 'Red
    lngPopArt(5) = 10040319 'Magenta
    lngPopArt(6) = 16653825 'Blue
    lngPopArt(7) = 16776960 'Cyan
    
    For i = 1 To 7
        If lngPopArt(i) >= inputColour Then
        
            If Abs(inputColour - lngPopArt(i)) < Abs(inputColour - lngPopArt(i - 1)) Then _
                PopArtColour = lngPopArt(i) Else PopArtColour = lngPopArt(i - 1)

            Exit For
            
            PopArtColour = lngPopArt(7)
        End If
    Next
    
    If PopArtColour = 0 Then PopArtColour = lngPopArt(7)
End Function

'Function that returns the nearest DarkMode colour
'to a given input colour
Public Function DarkModeColour(ByVal inputColour As Long) As Long
    Dim i As Integer
    Dim lngDarkMode(0 To 20) As Long 'Array mit den Farben

    lngDarkMode(0) = 5723136 '(Dummy)
    lngDarkMode(1) = 5723136
    lngDarkMode(2) = 7631617
    lngDarkMode(3) = 7864345
    lngDarkMode(4) = 8816385
    lngDarkMode(5) = 9606401
    lngDarkMode(6) = 9830439
    lngDarkMode(7) = 10068481
    lngDarkMode(8) = 10924800
    lngDarkMode(9) = 11730999
    lngDarkMode(10) = 11846656
    lngDarkMode(11) = 12966404
    lngDarkMode(12) = 13697099
    lngDarkMode(13) = 14610288
    lngDarkMode(14) = 15597666
    lngDarkMode(15) = 16056264
    lngDarkMode(16) = 16536990
    lngDarkMode(17) = 16549563
    lngDarkMode(18) = 16589439
    lngDarkMode(19) = 16627671
    lngDarkMode(20) = 16705522
    
    For i = 1 To 20
        If lngDarkMode(i) >= inputColour Then
        
            If Abs(inputColour - lngDarkMode(i)) < Abs(inputColour - lngDarkMode(i - 1)) Then _
                DarkModeColour = lngDarkMode(i) Else DarkModeColour = lngDarkMode(i - 1)

            Exit For
            
            DarkModeColour = lngDarkMode(20)
        End If
    Next
    
    If DarkModeColour = 0 Then DarkModeColour = lngDarkMode(20)
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

Public Sub GetColours_colMode()
    'Override settings dependent on the colour mode
    Select Case g_strColourMode
        Case "DARKMODE"
            objOption.COL_BACK = vbBlack
            objOption.COL_TEXT = vbWhite
            objOption.COL_RANKINGS = vbBlack
        Case "24H"
            objOption.DAYLIGHT = frmColourMode.scr24h.Value
            objOption.COL_BACK = objOption.DAYLIGHT_COL
            objOption.COL_TEXT = vbBlack
            objOption.COL_RANKINGS = objOption.DAYLIGHT_COL
        Case Else
            objOption.COL_BACK = xlNone
            objOption.COL_TEXT = vbBlack
            objOption.COL_RANKINGS = vbWhite
    End Select
End Sub

Sub GetColours_specRace()
    'Override settings dependent on the selected race
    Select Case objRace.RACE_ID
        Case "SPACE"
            objOption.COL_BACK = vbBlack
            objOption.COL_TEXT = vbWhite
            objOption.COL_RANKINGS = vbBlack
        Case "WATT18"
            objOption.COL_BACK = 11573124
            objOption.COL_TEXT = vbBlack
'        Case Else
'            objOption.COL_BACK = xlNone
'            objOption.COL_TEXT = vbBlack
    End Select
End Sub
