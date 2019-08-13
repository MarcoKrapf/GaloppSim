VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStart 
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11115
   OleObjectBlob   =   "frmStart.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up when a new race is started
'   UserForm frmStart

Dim i As Integer

Private Sub UserForm_Initialize()

    'Hide or show the "Race specific settings" button dependent on the chosen race
    cmdS4.Visible = False
    Select Case objRace.TRACK_SURFACE
        Case "M" 'Mudflats race
            With cmdS4
                .Visible = True
                .caption = GetText(g_arr_Text, "RACESPEC001")
            End With
        Case "MOON", "MARS", "JUPITER", "PLUTO", "SATURN" 'Space race
            With cmdS4
                .Visible = True
                .caption = GetText(g_arr_Text, "RACESPEC001")
            End With
    End Select
    
    'Hide the speed section if it should not be displayed
    Select Case objRace.RACE_ID
        Case "SPACE"
            cmdS5.Visible = False
    End Select
    
    'Hide the betting section if the betting mode is not enabled or betting is not allowed
    If objOption.BET_MODE = False Or objRace.BETTING_ALLOWED = "N" Then
        With Me
            .lblS8.Visible = False 'Label "Betting"
            .cmdS1.Visible = False 'Button "Add bet slip"
            .cmdS6.Visible = False 'Button "Odds (fixed rate)"
            .lstBetSlips.Visible = False 'ListBox with the bet slips
        End With
    End If
    
    'Format and hide the ListBox with the bet slips as long as there is no bet placed
    With lstBetSlips
        .Font = "Courier New"
        .Height = 76
        .Visible = False
    End With
    
    'Show the "Horse in focus" DropDown if the Focused Run mode is activated
    With Me
       .lblFocus.Visible = objOption.FOCUSED_RUN 'Label visible = TRUE or FALSE
       'Horse preview (8 segments from left(1) to right(8))
       .lblFocusColor1.Visible = objOption.FOCUSED_RUN 'Segment visible = TRUE or FALSE
       .lblFocusColor1.caption = "" 'no text
       .lblFocusColor2.Visible = objOption.FOCUSED_RUN
       .lblFocusColor2.caption = "" 'no text
       .lblFocusColor3.Visible = objOption.FOCUSED_RUN
       .lblFocusColor3.caption = "" 'no text
       .lblFocusColor4.Visible = objOption.FOCUSED_RUN
       .lblFocusColor4.caption = "" 'no text
       .lblFocusColor5.Visible = objOption.FOCUSED_RUN
       .lblFocusColor5.caption = "" 'no text
       .lblFocusColor6.Visible = objOption.FOCUSED_RUN
       .lblFocusColor6.caption = "" 'no text
       .lblFocusColor7.Visible = objOption.FOCUSED_RUN
       .lblFocusColor7.caption = "" 'no text
       .lblFocusColor8.Visible = objOption.FOCUSED_RUN
       .lblFocusColor8.caption = "" 'no text
       .lblFocusBorder.Visible = False
       With .cboFocus 'DropDown with the horses
            .Clear
            .Visible = objOption.FOCUSED_RUN 'ComboBox visible = True Or False
            .ColumnCount = 2 'Provide two columns for the horse name and number
            .ColumnWidths = "120 Pt;0 Pt" '120 pixels for the name, 0 for the number
            .Style = fmStyleDropDownList 'Allow only values from the item list, no free entries
        End With
    End With

    'Populate the "Horse in focus" DropDown with horses
    If objOption.FOCUSED_RUN Then
        objOption.FOCUSED_NR = 0 'Reset the focused horse
        For i = 2 To g_wksRaceData.Cells(rows.count, 6).End(xlUp).row 'Find the last row in the STATUS column of the race sheet
            If g_wksRaceData.Cells(i, 6).Value = "START" Then 'Consider only horses that are starting
                With cboFocus 'DropDown with the horses
                    .AddItem
                    .List(.ListCount - 1, 0) = _
                        g_wksRaceData.Cells(i, 7).Value & " (#" & _
                        g_wksRaceData.Cells(i, 5).Value & ")" 'Column 1 (visible): Compile the entry with the horse name and number (for displaying)
                    .List(.ListCount - 1, 1) = _
                        g_wksRaceData.Cells(i, 5).Value 'Column 2 (invisible): Horse number (for technical usage)
                End With
            End If
        Next i
    End If
    
    With cmdS2
        .SetFocus 'Set the focus on the "Start the race" button
    End With
    
    'Display the UserForm in the center of the Window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
    
End Sub

'Click on "Add new bet slip"
Private Sub cmdS1_Click()
    Call Gambler
End Sub

'Click on "Race specific settings"
Private Sub cmdS4_Click()
    Select Case objRace.TRACK_SURFACE
        Case "M" 'Mudflats race
            With frmTrackSettingsMudflats 'Pop-up
                .caption = objRace.TRACK_SURFACE_TEXT
                .show (vbModal)
            End With
        Case "MOON", "MARS", "JUPITER", "PLUTO", "SATURN" 'Space race
            With frmTracksSpaceRace 'Pop-up
                .caption = objRace.TRACK_SURFACE_TEXT
                .show (vbModal)
            End With
        Case Else
    End Select
End Sub

'Click on the ListBox with the bet slips
Private Sub lstBetSlips_Click()
    Dim i As Integer
    For i = 0 To lstBetSlips.ListCount - 1 'Loop through all items
        If lstBetSlips.SELECTED(i) Then 'Find the selected item
            Call ShowReceipt(i + 1) 'Show the receipt
            Exit For
        End If
    Next i
End Sub

'Click on "Start the race"
Private Sub cmdS2_Click()
    'In case of Focused Run mode: Check whether a horse in focus is selected
    If objOption.FOCUSED_RUN = True And objOption.FOCUSED_NR = 0 Then
        Call ShowInfoPopup(g_c_tool, _
            g_arr_Grammar(1) & " " & GetText(g_arr_Text, "START011") _
            , True, vbModal)
        Exit Sub
    End If
    
    objRace.STARTED = True 'Flag that indicates the race is started
    If g_colBetSlips.count = 0 Then objOption.BET_PLACED = False 'Flag that indicates whether bettings have been placed
    
    Unload Me
End Sub

'Click on "Speed and form"
Private Sub cmdS5_Click()
    Call ShowSpeed(False) 'Pop-up without odds
End Sub

'Click on "Odds (fixed rate)"
Private Sub cmdS6_Click()
    Call ShowSpeed(True) 'Pop-up with odds
End Sub

'Change of the focused horse
Private Sub cboFocus_Change()
    For i = 1 To UBound(g_arr_varHorses) 'Loop through all horses
        If CInt(cboFocus.List(cboFocus.ListIndex, 1)) = CInt(g_arr_varHorses(i, 11)) Then 'Compare the same data types (Integer)
            objOption.FOCUSED_NR = g_arr_varHorses(i, 11) 'Set the number of the focused horse
            
'            'Comment in the following lines for getting an understanding of the data types
'            Debug.Print TypeName(cboFocus.List(cboFocus.ListIndex, 1)) 'Value in the DropDown
'            Debug.Print TypeName(CInt(cboFocus.List(cboFocus.ListIndex, 1))) '... converted to Integer
'            Debug.Print TypeName(g_arr_varHorses(i, 11)) 'Value in the Array
'            Debug.Print TypeName(CInt(g_arr_varHorses(i, 11))) '... converted to Integer
            
            'Display the horse colour (8 segments)
            If IsArray(g_arr_varHorses(i, 2)) Then 'Multicoloured horse
                lblFocusColor1.BackColor = g_arr_varHorses(i, 2)(0)
                lblFocusColor2.BackColor = g_arr_varHorses(i, 2)(1)
                lblFocusColor3.BackColor = g_arr_varHorses(i, 2)(2)
                lblFocusColor4.BackColor = g_arr_varHorses(i, 2)(3)
                lblFocusColor5.BackColor = g_arr_varHorses(i, 2)(4)
                lblFocusColor6.BackColor = g_arr_varHorses(i, 2)(5)
                lblFocusColor7.BackColor = g_arr_varHorses(i, 2)(6)
                lblFocusColor8.BackColor = g_arr_varHorses(i, 2)(7)
            Else 'Monochrome horse
                lblFocusColor1.BackColor = g_arr_varHorses(i, 2)
                lblFocusColor2.BackColor = g_arr_varHorses(i, 2)
                lblFocusColor3.BackColor = g_arr_varHorses(i, 2)
                lblFocusColor4.BackColor = g_arr_varHorses(i, 2)
                lblFocusColor5.BackColor = g_arr_varHorses(i, 2)
                lblFocusColor6.BackColor = g_arr_varHorses(i, 2)
                lblFocusColor7.BackColor = g_arr_varHorses(i, 2)
                lblFocusColor8.BackColor = g_arr_varHorses(i, 2)
            End If
            Me.lblFocusBorder.Visible = True 'Show a border around the horse
            
            Exit For 'Leave the loop
        End If
    Next i
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then 'Click on "X" in the upper right corner of the UserForm
    
    'Pop-up with a security check before cancelling the race
    Call ShowMessagePopup(g_c_tool, GetText(g_arr_Text, "START004"), _
        enumButton.YesNo, vbModal)
        'Evaluate the return value
        If g_enumButton = enumButton.yes Then '>>Yes<< clicked
            objRace.STARTED = False
            If g_strPlayMode = "RS" Then Call RS_MenuAreaShow(False)
            If g_strPlayMode = "AI" Then Call AI_ExcelModeEnd
        Else '>>No<< clicked
            Cancel = 1 'Don´t close the pop-up
        End If
        
    Else 'Start button clicked
        objRace.STARTED = True
    End If
End Sub
