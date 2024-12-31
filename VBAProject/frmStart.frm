VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStart 
   ClientHeight    =   8292
   ClientLeft      =   -1980
   ClientTop       =   -8556
   ClientWidth     =   11916
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

Dim i As Integer, j As Integer

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
    Select Case objRace.SPECIAL
        Case "PARTICULATES" 'Particulates in the air
            With cmdS4
                .Visible = True
                .caption = GetText(g_arr_Text, "RACESPEC001")
            End With
        Case "DESERT" 'Desert race
            With cmdS4
                .Visible = True
                .caption = GetText(g_arr_Text, "RACESPEC001")
            End With
        Case "TUNNEL" 'Tunnel race
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
            .lblBet02.Visible = False 'Label with the number of betting slips
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
       .lblFocus.Visible = (objOption.FOCUSED_RUN = enumCamera.focus_horse) 'Label visibility
       'Horse preview (8 segments from left(1) to right(8))
       .lblFocusColour1.Visible = (objOption.FOCUSED_RUN = enumCamera.focus_horse) 'Segment visibility
       .lblFocusColour1.caption = "" 'no text
       .lblFocusColour2.Visible = (objOption.FOCUSED_RUN = enumCamera.focus_horse)
       .lblFocusColour2.caption = "" 'no text
       .lblFocusColour3.Visible = (objOption.FOCUSED_RUN = enumCamera.focus_horse)
       .lblFocusColour3.caption = "" 'no text
       .lblFocusColour4.Visible = (objOption.FOCUSED_RUN = enumCamera.focus_horse)
       .lblFocusColour4.caption = "" 'no text
       .lblFocusColour5.Visible = (objOption.FOCUSED_RUN = enumCamera.focus_horse)
       .lblFocusColour5.caption = "" 'no text
       .lblFocusColour6.Visible = (objOption.FOCUSED_RUN = enumCamera.focus_horse)
       .lblFocusColour6.caption = "" 'no text
       .lblFocusColour7.Visible = (objOption.FOCUSED_RUN = enumCamera.focus_horse)
       .lblFocusColour7.caption = "" 'no text
       .lblFocusColour8.Visible = (objOption.FOCUSED_RUN = enumCamera.focus_horse)
       .lblFocusColour8.caption = "" 'no text
       .lblFocusBorder.Visible = False
       With .cboFocus 'DropDown with the horses
            .Clear
            .Visible = (objOption.FOCUSED_RUN = enumCamera.focus_horse) 'ComboBox visibility
            .ColumnCount = 2 'Provide two columns for the horse name and number
            .ColumnWidths = "120 Pt;0 Pt" '120 pixels for the name, 0 for the number
            .Style = fmStyleDropDownList 'Allow only values from the item list, no free entries
        End With
    End With

    'Populate the "Horse in focus" DropDown with horses
    If objOption.FOCUSED_RUN = enumCamera.focus_horse Then
        objOption.FOCUSED_NR = 0 'Reset the focused horse
        For i = 1 To UBound(g_arr_varHorses())
            If g_arr_varHorses(i, 0) = "START" Then 'Consider only horses that are starting
                With cboFocus 'DropDown with the horses
                    .AddItem
                    .List(.ListCount - 1, 0) = _
                        g_arr_varHorses(i, 1) & " (#" & _
                        g_arr_varHorses(i, 11) & ")" 'Column 1 (visible): Compile the entry with the horse name and number (for displaying)
                    .List(.ListCount - 1, 1) = _
                        g_arr_varHorses(i, 11) 'Column 2 (invisible): Horse number (for technical usage)
                End With
            End If
        Next i
    End If
    
    'Show the Race Speed Monitor ListBox if RSMON is activated
    With Me
       .lblRSMon.Visible = (objOption.SPEEDMONITOR = True) 'Label visibility
       With .lstRSMon 'DropDown with the horses
            .Clear
            .Visible = (objOption.SPEEDMONITOR = True) 'ListBox visibility
            .ColumnCount = 2 'Provide two columns for the horse name and number
            .ColumnWidths = "120 Pt;0 Pt" '120 pixels for the name, 0 for the number
        End With
    End With
    
    'Populate the Race Speed Monitor ListBox with horses
    If objOption.SPEEDMONITOR Then
        For i = 1 To UBound(g_arr_varHorses())
            If g_arr_varHorses(i, 0) = "START" Then 'Consider only horses that are starting
                With lstRSMon 'DropDown with the horses
                    .AddItem
                    .List(.ListCount - 1, 0) = _
                        g_arr_varHorses(i, 1) & " (#" & _
                        g_arr_varHorses(i, 11) & ")" 'Column 1 (visible): Compile the entry with the horse name and number (for displaying)
                    .List(.ListCount - 1, 1) = _
                        g_arr_varHorses(i, 11) 'Column 2 (invisible): Horse number (for technical usage)
                End With
            End If
        Next i
        lstRSMon.MultiSelect = fmMultiSelectMulti
    End If
        
    'Hide the RSMon horse colour previews
    Call HideRSMonHorsePreview
    'Delete the label captions
    lbl_RSMon1_Col1.caption = ""
    lbl_RSMon1_Col2.caption = ""
    lbl_RSMon1_Col3.caption = ""
    lbl_RSMon1_Col4.caption = ""
    lbl_RSMon1_Col5.caption = ""
    lbl_RSMon1_Col6.caption = ""
    lbl_RSMon1_Col7.caption = ""
    lbl_RSMon1_Col8.caption = ""
    lbl_RSMon2_Col1.caption = ""
    lbl_RSMon2_Col2.caption = ""
    lbl_RSMon2_Col3.caption = ""
    lbl_RSMon2_Col4.caption = ""
    lbl_RSMon2_Col5.caption = ""
    lbl_RSMon2_Col6.caption = ""
    lbl_RSMon2_Col7.caption = ""
    lbl_RSMon2_Col8.caption = ""
    lbl_RSMon3_Col1.caption = ""
    lbl_RSMon3_Col2.caption = ""
    lbl_RSMon3_Col3.caption = ""
    lbl_RSMon3_Col4.caption = ""
    lbl_RSMon3_Col5.caption = ""
    lbl_RSMon3_Col6.caption = ""
    lbl_RSMon3_Col7.caption = ""
    lbl_RSMon3_Col8.caption = ""
    
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
            frmTrackSettingsMudflats.show (vbModal) 'Pop-up
        Case "MOON", "MARS", "JUPITER", "PLUTO", "SATURN" 'Space race
            frmTrackSettingsSpaceRace.show (vbModal) 'Pop-up
        Case Else
    End Select
    
    Select Case objRace.SPECIAL
        Case "PARTICULATES" 'Particulates in the air
            frmTrackSettingsAirQuality.show (vbModal) 'Pop-up
        Case "DESERT" 'Desert race
            frmTrackSettingsDesert.show (vbModal) 'Pop-up
        Case "TUNNEL" 'Tunnels
            frmTrackSettingsTunnels.show (vbModal) 'Pop-up
    End Select
    
End Sub

'Click on the ListBox with the bet slips
Private Sub lstBetSlips_Click()
    For i = 0 To lstBetSlips.ListCount - 1 'Loop through all items
        If lstBetSlips.SELECTED(i) Then 'Find the selected item
            Call ShowReceipt(i + 1) 'Show the receipt
            Exit For
        End If
    Next i
End Sub

'Click on the RSMon ListBox
Private Sub lstRSMon_Change()
    Dim count As Integer
    
    'Count the number of selected horses
    For i = 0 To lstRSMon.ListCount - 1
        If lstRSMon.SELECTED(i) Then count = count + 1
    Next i
    
    If count > 3 Then
        lstRSMon.SELECTED(lstRSMon.ListIndex) = False
    Else
        'Reset the RSMon collection
        Set g_colRSMon = Nothing
        Set g_colRSMon = New Collection
        
        For i = 0 To lstRSMon.ListCount - 1
            If lstRSMon.SELECTED(i) Then g_colRSMon.Add lstRSMon.List(i, 1)
        Next i
    End If
    
    'Hide the RSMon horse colour previews
    Call HideRSMonHorsePreview

    For i = 1 To g_colRSMon.count
        For j = 1 To UBound(g_arr_varHorses)
            If CDbl(g_arr_varHorses(j, 11)) = CDbl(g_colRSMon(i)) Then
                
                If g_colRSMon.item(1) = CDbl(g_colRSMon(i)) Then 'First item in the Collection
                    'Show the labels
                    lbl_RSMon1_Border.Visible = True
                    lbl_RSMon1_Col1.Visible = True
                    lbl_RSMon1_Col2.Visible = True
                    lbl_RSMon1_Col3.Visible = True
                    lbl_RSMon1_Col4.Visible = True
                    lbl_RSMon1_Col5.Visible = True
                    lbl_RSMon1_Col6.Visible = True
                    lbl_RSMon1_Col7.Visible = True
                    lbl_RSMon1_Col8.Visible = True
                    'Label colours
                    If IsArray(g_arr_varHorses(j, 2)) Then 'Multicoloured horse
                        lbl_RSMon1_Col1.BackColor = g_arr_varHorses(j, 2)(0)
                        lbl_RSMon1_Col2.BackColor = g_arr_varHorses(j, 2)(1)
                        lbl_RSMon1_Col3.BackColor = g_arr_varHorses(j, 2)(2)
                        lbl_RSMon1_Col4.BackColor = g_arr_varHorses(j, 2)(3)
                        lbl_RSMon1_Col5.BackColor = g_arr_varHorses(j, 2)(4)
                        lbl_RSMon1_Col6.BackColor = g_arr_varHorses(j, 2)(5)
                        lbl_RSMon1_Col7.BackColor = g_arr_varHorses(j, 2)(6)
                        lbl_RSMon1_Col8.BackColor = g_arr_varHorses(j, 2)(7)
                    Else 'Monochrome horse
                        lbl_RSMon1_Col1.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon1_Col2.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon1_Col3.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon1_Col4.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon1_Col5.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon1_Col6.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon1_Col7.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon1_Col8.BackColor = g_arr_varHorses(j, 2)
                    End If
                
                ElseIf g_colRSMon.item(2) = CDbl(g_colRSMon(i)) Then 'Second item in the Collection
                    'Show the labels
                    lbl_RSMon2_Border.Visible = True
                    lbl_RSMon2_Col1.Visible = True
                    lbl_RSMon2_Col2.Visible = True
                    lbl_RSMon2_Col3.Visible = True
                    lbl_RSMon2_Col4.Visible = True
                    lbl_RSMon2_Col5.Visible = True
                    lbl_RSMon2_Col6.Visible = True
                    lbl_RSMon2_Col7.Visible = True
                    lbl_RSMon2_Col8.Visible = True
                    'Label colours
                    If IsArray(g_arr_varHorses(j, 2)) Then 'Multicoloured horse
                        lbl_RSMon2_Col1.BackColor = g_arr_varHorses(j, 2)(0)
                        lbl_RSMon2_Col2.BackColor = g_arr_varHorses(j, 2)(1)
                        lbl_RSMon2_Col3.BackColor = g_arr_varHorses(j, 2)(2)
                        lbl_RSMon2_Col4.BackColor = g_arr_varHorses(j, 2)(3)
                        lbl_RSMon2_Col5.BackColor = g_arr_varHorses(j, 2)(4)
                        lbl_RSMon2_Col6.BackColor = g_arr_varHorses(j, 2)(5)
                        lbl_RSMon2_Col7.BackColor = g_arr_varHorses(j, 2)(6)
                        lbl_RSMon2_Col8.BackColor = g_arr_varHorses(j, 2)(7)
                    Else 'Monochrome horse
                        lbl_RSMon2_Col1.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon2_Col2.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon2_Col3.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon2_Col4.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon2_Col5.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon2_Col6.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon2_Col7.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon2_Col8.BackColor = g_arr_varHorses(j, 2)
                    End If
                    
                ElseIf g_colRSMon.item(3) = CDbl(g_colRSMon(i)) Then 'Third item in the Collection
                    'Show the labels
                    lbl_RSMon3_Border.Visible = True
                    lbl_RSMon3_Col1.Visible = True
                    lbl_RSMon3_Col2.Visible = True
                    lbl_RSMon3_Col3.Visible = True
                    lbl_RSMon3_Col4.Visible = True
                    lbl_RSMon3_Col5.Visible = True
                    lbl_RSMon3_Col6.Visible = True
                    lbl_RSMon3_Col7.Visible = True
                    lbl_RSMon3_Col8.Visible = True
                    'Label colours
                    If IsArray(g_arr_varHorses(j, 2)) Then 'Multicoloured horse
                        lbl_RSMon3_Col1.BackColor = g_arr_varHorses(j, 2)(0)
                        lbl_RSMon3_Col2.BackColor = g_arr_varHorses(j, 2)(1)
                        lbl_RSMon3_Col3.BackColor = g_arr_varHorses(j, 2)(2)
                        lbl_RSMon3_Col4.BackColor = g_arr_varHorses(j, 2)(3)
                        lbl_RSMon3_Col5.BackColor = g_arr_varHorses(j, 2)(4)
                        lbl_RSMon3_Col6.BackColor = g_arr_varHorses(j, 2)(5)
                        lbl_RSMon3_Col7.BackColor = g_arr_varHorses(j, 2)(6)
                        lbl_RSMon3_Col8.BackColor = g_arr_varHorses(j, 2)(7)
                    Else 'Monochrome horse
                        lbl_RSMon3_Col1.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon3_Col2.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon3_Col3.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon3_Col4.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon3_Col5.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon3_Col6.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon3_Col7.BackColor = g_arr_varHorses(j, 2)
                        lbl_RSMon3_Col8.BackColor = g_arr_varHorses(j, 2)
                    End If
                End If
            End If
        Next j
    Next i
    
End Sub

Private Sub HideRSMonHorsePreview()
    'Horse 1
    lbl_RSMon1_Border.Visible = False
    lbl_RSMon1_Col1.Visible = False
    lbl_RSMon1_Col2.Visible = False
    lbl_RSMon1_Col3.Visible = False
    lbl_RSMon1_Col4.Visible = False
    lbl_RSMon1_Col5.Visible = False
    lbl_RSMon1_Col6.Visible = False
    lbl_RSMon1_Col7.Visible = False
    lbl_RSMon1_Col8.Visible = False
    'Horse 2
    lbl_RSMon2_Border.Visible = False
    lbl_RSMon2_Col1.Visible = False
    lbl_RSMon2_Col2.Visible = False
    lbl_RSMon2_Col3.Visible = False
    lbl_RSMon2_Col4.Visible = False
    lbl_RSMon2_Col5.Visible = False
    lbl_RSMon2_Col6.Visible = False
    lbl_RSMon2_Col7.Visible = False
    lbl_RSMon2_Col8.Visible = False
    'Horse 3
    lbl_RSMon3_Border.Visible = False
    lbl_RSMon3_Col1.Visible = False
    lbl_RSMon3_Col2.Visible = False
    lbl_RSMon3_Col3.Visible = False
    lbl_RSMon3_Col4.Visible = False
    lbl_RSMon3_Col5.Visible = False
    lbl_RSMon3_Col6.Visible = False
    lbl_RSMon3_Col7.Visible = False
    lbl_RSMon3_Col8.Visible = False
End Sub

'Click on "Start the race"
Private Sub cmdS2_Click()
    'In case of Focused Run mode: Check whether a horse in focus is selected
    If objOption.FOCUSED_RUN = enumCamera.focus_horse And objOption.FOCUSED_NR = 0 Then
        Call ShowInfoPopup(g_c_tool, _
            GetText(g_arr_Text, "START017") & " " & g_arr_Grammar(2) & " " & GetText(g_arr_Text, "START011") _
            , True, vbModal, 24)
        Exit Sub
    End If
    
    'In case of RSMon: Check whether at least one horse is selected
    If objOption.SPEEDMONITOR Then
        If g_colRSMon.count = 0 Then
            Call ShowInfoPopup(g_c_tool, GetText(g_arr_Text, "START017") & " " & g_arr_Grammar(2) _
                & " " & GetText(g_arr_Text, "START013"), True, vbModal, 24)
            Exit Sub
        End If
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
                lblFocusColour1.BackColor = g_arr_varHorses(i, 2)(0)
                lblFocusColour2.BackColor = g_arr_varHorses(i, 2)(1)
                lblFocusColour3.BackColor = g_arr_varHorses(i, 2)(2)
                lblFocusColour4.BackColor = g_arr_varHorses(i, 2)(3)
                lblFocusColour5.BackColor = g_arr_varHorses(i, 2)(4)
                lblFocusColour6.BackColor = g_arr_varHorses(i, 2)(5)
                lblFocusColour7.BackColor = g_arr_varHorses(i, 2)(6)
                lblFocusColour8.BackColor = g_arr_varHorses(i, 2)(7)
            Else 'Monochrome horse
                lblFocusColour1.BackColor = g_arr_varHorses(i, 2)
                lblFocusColour2.BackColor = g_arr_varHorses(i, 2)
                lblFocusColour3.BackColor = g_arr_varHorses(i, 2)
                lblFocusColour4.BackColor = g_arr_varHorses(i, 2)
                lblFocusColour5.BackColor = g_arr_varHorses(i, 2)
                lblFocusColour6.BackColor = g_arr_varHorses(i, 2)
                lblFocusColour7.BackColor = g_arr_varHorses(i, 2)
                lblFocusColour8.BackColor = g_arr_varHorses(i, 2)
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
            If g_strPlayMode = "RS" Then Call RS_MenuAreaShow
            If g_strPlayMode = "AI" Then Call AI_ExcelModeEnd
        Else '>>No<< clicked
            Cancel = 1 'Don´t close the pop-up
        End If
        
    Else 'Start button clicked
        objRace.STARTED = True
    End If
End Sub
