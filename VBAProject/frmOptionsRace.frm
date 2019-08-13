VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionsRace 
   Caption         =   "[Race options]"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13455
   OleObjectBlob   =   "frmOptionsRace.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmOptionsRace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up for setting the race options
'   UserForm frmOptionsRace

Private Sub UserForm_Initialize()

    'Get the recommended zoom level
        objOption.ZOOM_LEVEL = basMainCode.ZoomLevelRecommendation
        
    'Draw the horse size preview
        lblscrRS_ro01.width = basMainCode.HorseSizePreview(objOption.ZOOM_LEVEL)(0)
        lblscrRS_ro01.Height = basMainCode.HorseSizePreview(objOption.ZOOM_LEVEL)(1)
    
    'Get the race participants
        Call basMainCode.GetAnimalGrammar
    
    'Settings of the sliders
        'Zoom level slider
        With scrRS_ro01
            .min = 1 'Minumum value
            .max = 3 'Maximum value
            .SmallChange = 1 'Value change when using the arrows
            .LargeChange = 1 'Value change when clicking inside the slider
        End With
        'Metres
        With scrRS_ro02
            .min = 50 'Minumum value
            .max = 1000 'Maximum value
            .SmallChange = 50 'value change when using the arrows
            .LargeChange = 50 'value change when clicking inside the slider
        End With
        'Race speed factor
        With scrRS_ro03
            .min = 1 'Minumum value
            .max = 5 'Maximum value
            .SmallChange = 1 'value change when using the arrows
            .LargeChange = 1 'value change when clicking inside the slider
        End With
        
    'Settings of the toggle buttons
        'Toggle button for the colours of the photo of the finish
         togRS_ro01 = objOption.PHOTO_BW
         Call togRS_ro01_Click
    
    With Me
        'Captions
        .caption = GetText(g_arr_Text, "USERFORM001")
        .lblRS_ro01.caption = GetText(g_arr_Text, "RACEOPT001")
        .lblRS_ro05.caption = GetText(g_arr_Text, "START005")
        .fraRS_ro01.caption = GetText(g_arr_Text, "RACEOPT009")
        .optRS_ro01.caption = GetText(g_arr_Text, "RACEOPT010")
        .optRS_ro02.caption = GetText(g_arr_Text, "RACEOPT011")
        .chkRS_ro02.caption = GetText(g_arr_Text, "RACEOPT017")
        .chkRS_ro08.caption = GetText(g_arr_Text, "RACEOPT015")
        .chkRS_ro09.caption = GetText(g_arr_Text, "RACEOPT016")
        .lblRS_ro06.caption = GetText(g_arr_Text, "RACEOPT013")
        .fraRS_ro05.caption = GetText(g_arr_Text, "RACEOPT013")
        .chkRS_ro10.caption = GetText(g_arr_Text, "RACEOPT014")
        .chkRS_ro13.caption = GetText(g_arr_Text, "RACEOPT032")
        .chkRS_ro14.caption = GetText(g_arr_Text, "RACEOPT034")
        .chkRS_ro11a.caption = GetText(g_arr_Text, "RACEOPT022")
        .chkRS_ro11b.caption = GetText(g_arr_Text, "RACEOPT024")
        .fraRS_ro03.caption = GetText(g_arr_Text, "RACEOPT018")
        .chkRS_ro12a.caption = GetText(g_arr_Text, "RACEOPT019")
        .chkRS_ro12b.caption = GetText(g_arr_Text, "RACEOPT020")
        .chkRS_ro12c.caption = GetText(g_arr_Text, "RACEOPT021")
        .fraRS_ro04.caption = GetText(g_arr_Text, "RACEOPT026")
        .chkRS_ro01a.caption = g_arr_Grammar(1) & " " & GetText(g_arr_Text, "START006") & " " & GetText(g_arr_Text, "RACEOPT027")
        .chkRS_ro01b.caption = GetText(g_arr_Text, "RACEOPT004")
        .cmdRS_ro01.caption = GetText(g_arr_Text, "BTN014")
        .lblRS_ro08.caption = GetText(g_arr_Text, "RACEOPT035")
        .fraRS_ro02.caption = GetText(g_arr_Text, "ZOOM001") & ": " & ZoomLevelText(objOption.ZOOM_LEVEL)
        .lblscrRS_ro02.caption = GetText(g_arr_Text, "ZOOM006") & " " & g_arr_Grammar(5)
        .fraRS_ro07.caption = GetText(g_arr_Text, "RACEOPT036") & " " _
                                & GetText(g_arr_Text, "RACE020") & " " & scrRS_ro02.Value & GetText(g_arr_Text, "RACE008")
        .cmdRS_ro02a.caption = GetText(g_arr_Text, "RACEOPT029")
        .cmdRS_ro02b.caption = GetText(g_arr_Text, "RACEOPT030")
        .cmdRS_ro02c.caption = GetText(g_arr_Text, "RACEOPT031")
        .lblRS_ro07.caption = GetText(g_arr_Text, "RACEOPT033")
        .fraRS_ro08.caption = GetText(g_arr_Text, "RACEOPT002")
        .chkRS_ro04.caption = GetText(g_arr_Text, "RACEOPT005")
        .chkRS_ro05.caption = GetText(g_arr_Text, "RACEOPT006")
        .chkRS_ro06.caption = GetText(g_arr_Text, "RACEOPT007")
        .chkRS_ro07.caption = GetText(g_arr_Text, "RACEOPT008")
        .fraRS_ro09.caption = GetText(g_arr_Text, "RACEOPT003")
        .lblRS_ro09.caption = GetText(g_arr_Text, "RACEOPT041")
        .cmdRS_ro03a.caption = GetText(g_arr_Text, "RACEOPT037")
        .cmdRS_ro03b.caption = GetText(g_arr_Text, "RACEOPT038")
        .opt_ro01.caption = GetText(g_arr_Text, "RACEOPT039")
        .opt_ro02.caption = GetText(g_arr_Text, "RACEOPT040")
        .lblRS_ro10.caption = GetText(g_arr_Text, "RACEOPT042")
        .lblRS_ro11.caption = objOption.SPEED_FACTOR & GetText(g_arr_Text, "RACEOPT043")
        .lblRS_ro12a.caption = GetText(g_arr_Text, "RACEOPT046")
        .lblRS_ro12b.caption = GetText(g_arr_Text, "RACEOPT047")
        .fraRS_ro12.caption = GetText(g_arr_Text, "RACEOPT044")
        .chkRS_ro15.caption = g_arr_Grammar(1) & " " & GetText(g_arr_Text, "RACEOPT045")
        .chkRS_ro16.caption = GetText(g_arr_Text, "RACEOPT048")
        .chkRS_ro17.caption = GetText(g_arr_Text, "RACEOPT049a") & " " & g_arr_Grammar(4) & " " & GetText(g_arr_Text, "RACEOPT049b") & "."
        .fraRS_ro10.caption = GetText(g_arr_Text, "BTN004")
        .fraRS_ro13.caption = GetText(g_arr_Text, "RACEOPT052")
        .chkRS_ro18.caption = GetText(g_arr_Text, "RACEOPT053")
        
        'ControlTipTexts
        .optRS_ro01.ControlTipText = GetText(g_arr_Text, "TIP001a") & " " & g_arr_Grammar(4) & " " & GetText(g_arr_Text, "TIP001b")
        .optRS_ro02.ControlTipText = GetText(g_arr_Text, "TIP001a") & " " & g_arr_Grammar(4) & " " & GetText(g_arr_Text, "TIP002")
        .chkRS_ro01a.ControlTipText = GetText(g_arr_Text, "TIP003")
        .chkRS_ro02.ControlTipText = GetText(g_arr_Text, "TIP004a") & " " & g_arr_Grammar(8) & " " & GetText(g_arr_Text, "TIP004b")
        .chkRS_ro04.ControlTipText = GetText(g_arr_Text, "TIP005a") & " " & g_arr_Grammar(6) & " " & GetText(g_arr_Text, "TIP005b")
        .chkRS_ro05.ControlTipText = GetText(g_arr_Text, "TIP006") & " " & g_arr_Grammar(6) & " " & GetText(g_arr_Text, "TIP005b")
        .chkRS_ro06.ControlTipText = GetText(g_arr_Text, "TIP007")
        .chkRS_ro07.ControlTipText = GetText(g_arr_Text, "TIP008a") & " " & g_arr_Grammar(6) & " " & GetText(g_arr_Text, "TIP008b")
        .chkRS_ro08.ControlTipText = GetText(g_arr_Text, "TIP009a") & " " & g_arr_Grammar(6) & " " & GetText(g_arr_Text, "TIP009b")
        .chkRS_ro09.ControlTipText = GetText(g_arr_Text, "TIP010")
        .chkRS_ro10.ControlTipText = GetText(g_arr_Text, "TIP011")
        .chkRS_ro11a.ControlTipText = GetText(g_arr_Text, "RACEOPT023")
        .chkRS_ro11b.ControlTipText = GetText(g_arr_Text, "RACEOPT025")
        .chkRS_ro15.ControlTipText = g_arr_Grammar(3) & " " & GetText(g_arr_Text, "TIP016")
        .chkRS_ro12a.ControlTipText = GetText(g_arr_Text, "TIP001a") & " " & g_arr_Grammar(4) & " " & GetText(g_arr_Text, "TIP017a") & " " & g_arr_Grammar(6) & " " & GetText(g_arr_Text, "TIP017b")
        .chkRS_ro12b.ControlTipText = GetText(g_arr_Text, "TIP018") & " " & g_arr_Grammar(4)
        .chkRS_ro12c.ControlTipText = GetText(g_arr_Text, "TIP019")
        .chkRS_ro16.ControlTipText = GetText(g_arr_Text, "TIP020")
        .scrRS_ro03.ControlTipText = GetText(g_arr_Text, "TIP021a") & " " & g_arr_Grammar(4) & " " & GetText(g_arr_Text, "TIP021b")
        .cmdRS_ro03a.ControlTipText = GetText(g_arr_Text, "TIP022")
        .cmdRS_ro03b.ControlTipText = GetText(g_arr_Text, "TIP023")
        .chkRS_ro18.ControlTipText = GetText(g_arr_Text, "TIP028")
        
        'Get values
        .optRS_ro01.Value = (objOption.TACTICS = False)
        .optRS_ro02.Value = (objOption.TACTICS = True)
        .chkRS_ro01a.Value = objOption.FOCUSED_RUN
        .chkRS_ro01b.Value = objOption.HIGHLIGHT_FOC
        .chkRS_ro01b.Enabled = (objOption.FOCUSED_RUN = True) 'Enabled only if the "Focused Run" checkbox is ticked
        .chkRS_ro02.Value = objOption.HOOFPRINTS
        .scrRS_ro01.Value = objOption.ZOOM_LEVEL
        .scrRS_ro02.Value = objOption.METRES_DISPLAY
        .chkRS_ro04.Value = objOption.NAMES_LEFT
        .chkRS_ro05.Value = objOption.COLOURS_LEFT
        .chkRS_ro06.Value = objOption.HIGHLIGHT_FAV
        .chkRS_ro07.Value = objOption.NAMES_FINISH
        .chkRS_ro17.Value = objOption.NAMES_PHOTO
        .chkRS_ro08.Value = objOption.RANKING_COL
        .chkRS_ro09.Value = objOption.RANKING_DELAY
        .chkRS_ro10.Value = objOption.RACE_INFO
        .opt_ro01.Value = objOption.RACE_INFO_POP
        .opt_ro02.Value = objOption.RACE_INFO_WKS
        .opt_ro01.Enabled = (objOption.RACE_INFO = True)  'Enabled only if the "Race info" checkbox is ticked
        .opt_ro02.Enabled = (objOption.RACE_INFO = True)  'Enabled only if the "Race info" checkbox is ticked
        .chkRS_ro13.Value = objOption.RACE_INFO_LEADER
        .chkRS_ro13.Enabled = (objOption.RACE_INFO = True)  'Enabled only if the "Race info" checkbox is ticked
        .chkRS_ro14.Value = objOption.RACE_INFO_PROGRESS
        .chkRS_ro14.Enabled = (objOption.RACE_INFO = True)  'Enabled only if the "Race info" checkbox is ticked
        cmdRS_ro02a.Enabled = (objOption.RACE_INFO = True)  'Enabled only if the "Race info" checkbox is ticked
        cmdRS_ro02b.Enabled = (objOption.RACE_INFO = True)  'Enabled only if the "Race info" checkbox is ticked
        cmdRS_ro02c.Enabled = (objOption.RACE_INFO = True)  'Enabled only if the "Race info" checkbox is ticked
        lblRS_ro07.BackColor = objOption.RACE_INFO_COL_B
        lblRS_ro07.ForeColor = objOption.RACE_INFO_COL_F
        lblRS_ro07.Visible = (objOption.RACE_INFO = True)  'Enabled only if the "Race info" checkbox is ticked
        .chkRS_ro11a.Value = objOption.BET_MODE
        .chkRS_ro11b.Value = objOption.BET_ANALYSIS
        .chkRS_ro11b.Enabled = (objOption.BET_MODE = True) 'Enabled only if the "Placing bets" checkbox is ticked
        .chkRS_ro12a.Value = objOption.SLIPSTREAM
        .chkRS_ro12b.Value = objOption.SLIPSTREAM_DBL
        .chkRS_ro12b.Enabled = (objOption.SLIPSTREAM = True) 'Enabled only if the "Slipstreaming" checkbox is ticked
        .chkRS_ro12c.Value = objOption.SLIPSTREAM_SHOW
        .chkRS_ro12c.Enabled = (objOption.SLIPSTREAM = True) 'Enabled only if the "Slipstreaming" checkbox is ticked
        .scrRS_ro03.Value = objOption.SPEED_FACTOR
        .chkRS_ro15.Value = objOption.REFUSE_RUN
        .chkRS_ro16.Value = objOption.SPEECH
        .chkRS_ro18.Value = objOption.MOMENTUM
    End With
    
    'Set the initial race info preview colours
    Call DefaultPreviewColours
    
    'Display the UserForm in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

'Option button "Race info in a pop-up"
Private Sub opt_ro01_Click()
    Call DefaultPreviewColours
End Sub

'Option button "Race info on the worksheet"
Private Sub opt_ro02_Click()
    Call DefaultPreviewColours
End Sub

'Default colours of the race info preview
Private Sub DefaultPreviewColours()
    If opt_ro01 Then 'Default in a pop-up: Black font colour on grey background
        If lblRS_ro07.BackColor = vbWhite Then lblRS_ro07.BackColor = -2147483633 'Grey
        If lblRS_ro07.ForeColor = vbBlack Then lblRS_ro07.ForeColor = -2147483630 'Black
    End If
    If opt_ro02 Then 'Default on the worksheet: Black font colour on white background
        If lblRS_ro07.BackColor = -2147483633 Then lblRS_ro07.BackColor = vbWhite
        If lblRS_ro07.ForeColor = -2147483630 Then lblRS_ro07.ForeColor = vbBlack
    End If
End Sub

'Button for the race info background colour
Private Sub cmdRS_ro02a_Click()
    lblRS_ro07.BackColor = basAuxiliary.ColPick(lblRS_ro07.BackColor)
End Sub

'Button for the race info font colour
Private Sub cmdRS_ro02b_Click()
    lblRS_ro07.ForeColor = basAuxiliary.ColPick(lblRS_ro07.ForeColor)
End Sub

'Button for resetting the race info colours
Private Sub cmdRS_ro02c_Click()
    If opt_ro01 Then 'Default in a pop-up: Black font colour on grey background
        'Take the colours from the reset button
        lblRS_ro07.BackColor = cmdRS_ro02c.BackColor 'Grey
        lblRS_ro07.ForeColor = cmdRS_ro02c.ForeColor 'Black
    End If
    If opt_ro02 Then 'Default on the worksheet: Black font colour on white background
        lblRS_ro07.BackColor = vbWhite
        lblRS_ro07.ForeColor = vbBlack
    End If
End Sub

'Click on the OK button
Private Sub cmdRS_ro01_Click()
    'Set the selected values
    Select Case True
        Case optRS_ro01.Value
            objOption.TACTICS = False
        Case optRS_ro02.Value
            objOption.TACTICS = True
    End Select
    objOption.SLIPSTREAM = chkRS_ro12a.Value
    objOption.SLIPSTREAM_DBL = chkRS_ro12b.Value
    objOption.SLIPSTREAM_SHOW = chkRS_ro12c.Value
    objOption.FOCUSED_RUN = chkRS_ro01a.Value
    objOption.HIGHLIGHT_FOC = chkRS_ro01b.Value
    objOption.HOOFPRINTS = chkRS_ro02.Value
    objOption.ZOOM_LEVEL = scrRS_ro01.Value
    objOption.METRES_DISPLAY = scrRS_ro02.Value
    objOption.NAMES_LEFT = chkRS_ro04.Value
    objOption.COLOURS_LEFT = chkRS_ro05.Value
    objOption.HIGHLIGHT_FAV = chkRS_ro06.Value
    objOption.NAMES_FINISH = chkRS_ro07.Value
    objOption.NAMES_PHOTO = chkRS_ro17.Value
    objOption.PHOTO_BW = togRS_ro01.Value
    objOption.RANKING_COL = chkRS_ro08.Value
    objOption.RANKING_DELAY = chkRS_ro09.Value
    objOption.RACE_INFO = chkRS_ro10.Value
    objOption.RACE_INFO_POP = opt_ro01.Value
    objOption.RACE_INFO_WKS = opt_ro02.Value
    objOption.RACE_INFO_LEADER = chkRS_ro13.Value
    objOption.RACE_INFO_PROGRESS = chkRS_ro14.Value
    objOption.RACE_INFO_COL_B = lblRS_ro07.BackColor
    objOption.RACE_INFO_COL_F = lblRS_ro07.ForeColor
    objOption.BET_MODE = chkRS_ro11a.Value
    objOption.BET_ANALYSIS = chkRS_ro11b.Value
    objOption.SPEED_FACTOR = scrRS_ro03.Value
    objOption.REFUSE_RUN = chkRS_ro15.Value
    objOption.SPEECH = chkRS_ro16.Value
    objOption.MOMENTUM = chkRS_ro18.Value
    
    'Adapt the caption of the race start button ("Betting and race" or "xxxxxxx")
    If g_strPlayMode = "RS" Then
        g_wksRace.OLEObjects("startrace").Object.caption = GetText(g_arr_Text, getCaptionStartBtn(objOption.BET_MODE))
    Else 'AI edition - refresh the ribbon
        g_RibbonGaloppSim.Invalidate
    End If
    
    'Close the UserForm
    Unload Me
End Sub

'Click on the "Focused Run" checkbox
Private Sub chkRS_ro01a_Click()
    Me.chkRS_ro01b.Enabled = (Me.chkRS_ro01a.Value = True) 'Set the status dependent on the "Focused Run" main checkbox
End Sub

'Click on the "Placing bets" checkbox
Private Sub chkRS_ro11a_Click()
    Me.chkRS_ro11b.Enabled = (Me.chkRS_ro11a.Value = True) 'Set the status dependent on the "Placing bets" main checkbox
End Sub

'Click on the "Slipstreaming" checkbox
Private Sub chkRS_ro12a_Click()
    With Me
        .chkRS_ro12b.Enabled = (.chkRS_ro12a.Value = True) 'Set the status dependent on the "Slipstreaming" main checkbox
        .chkRS_ro12c.Enabled = (.chkRS_ro12a.Value = True) 'Set the status dependent on the "Slipstreaming" main checkbox
    End With
End Sub

'Click on the toggle button for the colours on the photo of the finish
Private Sub togRS_ro01_Click()
    With togRS_ro01
        If .Value Then 'If pressed (value = true)
            .BackColor = &H8000000C 'Grey
            lblRS_ro13.caption = GetText(g_arr_Text, "RACEOPT051")
        Else 'If not pressed
            .BackColor = vbRed 'Red
            lblRS_ro13.caption = GetText(g_arr_Text, "RACEOPT050")
        End If
    End With
End Sub

'Click on the "Race info" checkbox
Private Sub chkRS_ro10_Click()
    'Set the status dependent on the "Race info" main checkbox
    With Me
        .opt_ro01.Enabled = (Me.chkRS_ro10.Value = True)
        .opt_ro02.Enabled = (Me.chkRS_ro10.Value = True)
        .chkRS_ro13.Enabled = (Me.chkRS_ro10.Value = True)
        .chkRS_ro14.Enabled = (Me.chkRS_ro10.Value = True)
        .cmdRS_ro02a.Enabled = (Me.chkRS_ro10.Value = True)
        .cmdRS_ro02b.Enabled = (Me.chkRS_ro10.Value = True)
        .cmdRS_ro02c.Enabled = (Me.chkRS_ro10.Value = True)
        .lblRS_ro07.Visible = (Me.chkRS_ro10.Value = True)
    End With
End Sub

'Change of the zoom level slider position
Private Sub scrRS_ro01_Change()
    'Adapt the horse size preview
    lblscrRS_ro01.width = basMainCode.HorseSizePreview(scrRS_ro01.Value)(0)
    lblscrRS_ro01.Height = basMainCode.HorseSizePreview(scrRS_ro01.Value)(1)
    'Adapt the zoom level text
        fraRS_ro02.caption = GetText(g_arr_Text, "ZOOM001") & ": " & basMainCode.ZoomLevelText(scrRS_ro01.Value)
End Sub

'Change of the race track metres slider position
Private Sub scrRS_ro02_Change()
    Dim markerdistance As String '[xxx]m, e.g. 250m
    Dim i As Integer, j As Integer
    For i = 1 To (4000 / scrRS_ro02.Value) 'Calculate the number of markers
        markerdistance = markerdistance & (i * scrRS_ro02.Value) & GetText(g_arr_Text, "RACE008")
        For j = 1 To (scrRS_ro02.Value / scrRS_ro02.min) 'Add space between the markers
            markerdistance = markerdistance & " "
        Next
    Next
    'Adapt the texts
    lblscrRS_ro03.caption = markerdistance 'Preview below the slider
    fraRS_ro07.caption = GetText(g_arr_Text, "RACEOPT036") & " " _
                        & GetText(g_arr_Text, "RACE020") & " " & scrRS_ro02.Value & GetText(g_arr_Text, "RACE008") 'Frame caption
End Sub

'Change of the race speed factor slider position
Private Sub scrRS_ro03_Change()
    'Adapt the race speed factor label
        lblRS_ro11.caption = scrRS_ro03.Value & GetText(g_arr_Text, "RACEOPT043")
End Sub

'Click on the "Save settings" button
Private Sub cmdRS_ro03a_Click()
    Call basInputOutput.SaveRaceOptions
End Sub

'Click on the "Load settings" button
Private Sub cmdRS_ro03b_Click()
    Call basInputOutput.LoadRaceOptions
End Sub
