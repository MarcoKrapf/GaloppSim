VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionsRace 
   Caption         =   "[Race options]"
   ClientHeight    =   8940
   ClientLeft      =   -1812
   ClientTop       =   -7860
   ClientWidth     =   17964
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
    
    'Get the race participants
        Call basMainCode.GetAnimalGrammar
    
    'Settings of the sliders
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
        'Spectators (in %)
        With scrSpec
            .min = 0 'Minumum value
            .max = 100 'Maximum value
            .SmallChange = 25 'value change when using the arrows
            .LargeChange = 25 'value change when clicking inside the slider
        End With
        'Momentum speed bars refresh rate
        With scrMomRefr
            .min = 1 'Minumum value
            .max = 60 'Maximum value
            .SmallChange = 1 'value change when using the arrows
            .LargeChange = 10 'value change when clicking inside the slider
        End With
        'Race speed monitor refresh rate
        With scrRSMonRefr
            .min = 10 'Minumum value
            .max = 50 'Maximum value
            .SmallChange = 1 'value change when using the arrows
            .LargeChange = 10 'value change when clicking inside the slider
        End With
        'Start refusal rate
        With scrRefusalRate
            .min = 10 'Minumum value
            .max = 1000 'Maximum value
            .SmallChange = 1 'value change when using the arrows
            .LargeChange = 10 'value change when clicking inside the slider
        End With
        'Slipstream effect
        With scrSlipstream
            .min = 0 'Minumum value
            .max = 2 'Maximum value
            .SmallChange = 1 'value change when using the arrows
            .LargeChange = 1 'value change when clicking inside the slider
        End With
        
    'Settings of the toggle buttons
        'Toggle button for the colours of the photo of the finish
         togRS_ro01 = objOption.PHOTO_BW
         Call togRS_ro01_Click
    
    With Me
        .Height = 475
        .width = 910
        .StartUpPosition = 0 'Place the UserForm in the upper left corner
        .top = 0
        .left = 0
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
        .fraRS_ro03.caption = GetText(g_arr_Text, "RACEOPT020")
        .lbl_slipstream_0.caption = GetText(g_arr_Text, "RACEOPT074")
        .lbl_slipstream_1.caption = GetText(g_arr_Text, "RACEOPT075")
        .lbl_slipstream_2.caption = GetText(g_arr_Text, "RACEOPT076")
        .chk_SlipstreamShow.caption = GetText(g_arr_Text, "RACEOPT021")
        .fraRS_ro04.caption = GetText(g_arr_Text, "RACEOPT026")
        .opt_foc01.caption = GetText(g_arr_Text, "RACEOPT056")
        .opt_foc02.caption = GetText(g_arr_Text, "RACEOPT058") & " " & g_arr_Grammar(2) & " " & GetText(g_arr_Text, "START006") & " " & GetText(g_arr_Text, "RACEOPT027")
        .chk_foc02b.caption = GetText(g_arr_Text, "RACEOPT004")
        .opt_foc03.caption = GetText(g_arr_Text, "RACEOPT057")
        .cmdRS_ro01.caption = GetText(g_arr_Text, "BTN014")
        .lblRS_ro08.caption = GetText(g_arr_Text, "RACEOPT035")
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
        .chkFavAnn.caption = GetText(g_arr_Text, "RACEOPT090")
        .chkRS_ro07.caption = GetText(g_arr_Text, "RACEOPT008")
        .fraRS_ro09.caption = GetText(g_arr_Text, "RACEOPT003")
        .lblRS_ro09.caption = GetText(g_arr_Text, "RACEOPT041")
        .cmdRS_ro03a.caption = GetText(g_arr_Text, "RACEOPT037")
        .cmdRS_ro03b.caption = GetText(g_arr_Text, "RACEOPT038")
        .opt_ro01.caption = GetText(g_arr_Text, "RACEOPT039")
        .opt_ro02.caption = GetText(g_arr_Text, "RACEOPT040")
        .fraRS_ro11.caption = GetText(g_arr_Text, "RACEOPT042")
        .lblRS_ro11.caption = objOption.SPEED_FACTOR & GetText(g_arr_Text, "RACEOPT043")
        .lblRS_ro12a.caption = GetText(g_arr_Text, "RACEOPT046")
        .lblRS_ro12b.caption = GetText(g_arr_Text, "RACEOPT047")
        .fraRS_ro12.caption = GetText(g_arr_Text, "RACEOPT044")
        .chkRS_ro15.caption = g_arr_Grammar(3) & " " & GetText(g_arr_Text, "RACEOPT045")
        .chkRS_ro16.caption = GetText(g_arr_Text, "RACEOPT048")
        .chkRS_ro17.caption = GetText(g_arr_Text, "RACEOPT049a") & " " & g_arr_Grammar(4) & " " & GetText(g_arr_Text, "RACEOPT049b")
        .fraRS_ro10.caption = GetText(g_arr_Text, "BTN004")
        .fraRS_ro13.caption = GetText(g_arr_Text, "RACEOPT052")
        .chk_MOM1.caption = GetText(g_arr_Text, "RACEOPT053")
        .chk_MOM2.caption = GetText(g_arr_Text, "RACEOPT077")
        .fraRSMon.caption = GetText(g_arr_Text, "RACEOPT079")
        .chkRSMon.caption = GetText(g_arr_Text, "RACEOPT080")
        .optRSMon1.caption = GetText(g_arr_Text, "RACEOPT092")
        .optRSMon2.caption = GetText(g_arr_Text, "RACEOPT093")
        .chkTribunes.caption = GetText(g_arr_Text, "RACEOPT059")
        .fraRS_ro14.caption = GetText(g_arr_Text, "RACEOPT060")
        .lblSpec1.caption = GetText(g_arr_Text, "RACEOPT061")
        .lblSpec2.caption = GetText(g_arr_Text, "RACEOPT062")
        .lblMomRefr.caption = GetText(g_arr_Text, "RACEOPT087")
        .lblRefrLow1.caption = GetText(g_arr_Text, "RACEOPT088")
        .lblRefrHigh1.caption = GetText(g_arr_Text, "RACEOPT089")
        .lblRSMonRefr.caption = GetText(g_arr_Text, "RACEOPT087")
        .lblRefrLow2.caption = GetText(g_arr_Text, "RACEOPT088")
        .lblRefrHigh2.caption = GetText(g_arr_Text, "RACEOPT089")
        .lblRefusalRate.caption = "1 " & GetText(g_arr_Text, "RACEOPT065") & " " & objOption.REFUSAL_RATE & " " & GetText(g_arr_Text, "RACEOPT066")
        .chk_TEO1.caption = GetText(g_arr_Text, "RACEOPT068")
        .chk_TEO2.caption = GetText(g_arr_Text, "RACEOPT069")
        .lblRS_ro14.caption = GetText(g_arr_Text, "RACEOPT071")
        .lblRS_ro15.caption = GetText(g_arr_Text, "RACEOPT072")
        .chkRS_ro19.caption = GetText(g_arr_Text, "RACEOPT073")
        .chkAutoSaveRace.caption = GetText(g_arr_Text, "RACEOPT091")
        .lblRS_ro16.caption = GetText(g_arr_Text, "RACEOPT094")
        .fraRS_ro16.caption = GetText(g_arr_Text, "RACEOPT095")
        .opt_grid1.caption = g_arr_Grammar(3) & " " & GetText(g_arr_Text, "RACEOPT096")
        .opt_grid2.caption = g_arr_Grammar(3) & " " & GetText(g_arr_Text, "RACEOPT097")
        
        'ControlTipTexts
        .optRS_ro01.ControlTipText = GetText(g_arr_Text, "TIP001a") & " " & g_arr_Grammar(4) & " " & GetText(g_arr_Text, "TIP001b")
        .optRS_ro02.ControlTipText = GetText(g_arr_Text, "TIP001a") & " " & g_arr_Grammar(4) & " " & GetText(g_arr_Text, "TIP002")
        .opt_foc01.ControlTipText = GetText(g_arr_Text, "TIP031")
        .opt_foc02.ControlTipText = GetText(g_arr_Text, "TIP003")
        .chk_foc02b.ControlTipText = GetText(g_arr_Text, "TIP032")
        .opt_foc03.ControlTipText = GetText(g_arr_Text, "TIP033")
        .chkRS_ro02.ControlTipText = GetText(g_arr_Text, "TIP004a") & " " & g_arr_Grammar(8) & " " & GetText(g_arr_Text, "TIP004b")
        .chkRS_ro04.ControlTipText = GetText(g_arr_Text, "TIP005a") & " " & g_arr_Grammar(6) & " " & GetText(g_arr_Text, "TIP005b")
        .chkRS_ro05.ControlTipText = GetText(g_arr_Text, "TIP006") & " " & g_arr_Grammar(6) & " " & GetText(g_arr_Text, "TIP005b")
        .chkRS_ro06.ControlTipText = GetText(g_arr_Text, "TIP007")
        .chkFavAnn.ControlTipText = GetText(g_arr_Text, "TIP001a") & " " & g_arr_Grammar(4) & " " & GetText(g_arr_Text, "TIP055")
        .chkRS_ro07.ControlTipText = GetText(g_arr_Text, "TIP008a") & " " & g_arr_Grammar(6) & " " & GetText(g_arr_Text, "TIP008b")
        .chkRS_ro08.ControlTipText = GetText(g_arr_Text, "TIP009a") & " " & g_arr_Grammar(6) & " " & GetText(g_arr_Text, "TIP009b")
        .chkRS_ro09.ControlTipText = GetText(g_arr_Text, "TIP010")
        .chkRS_ro10.ControlTipText = GetText(g_arr_Text, "TIP011")
        .chkRS_ro11a.ControlTipText = GetText(g_arr_Text, "RACEOPT023")
        .chkRS_ro11b.ControlTipText = GetText(g_arr_Text, "RACEOPT025")
        .chkRS_ro15.ControlTipText = g_arr_Grammar(3) & " " & GetText(g_arr_Text, "TIP016")
        .fraRS_ro03.ControlTipText = GetText(g_arr_Text, "RACEOPT019")
        .scrSlipstream.ControlTipText = GetText(g_arr_Text, "TIP001a") & " " & g_arr_Grammar(4) & " " & GetText(g_arr_Text, "TIP017a") & " " & g_arr_Grammar(6) & " " & GetText(g_arr_Text, "TIP017b") & ". " _
             & GetText(g_arr_Text, "TIP018") & " " & g_arr_Grammar(4) & "."
        .lbl_slipstream_1.ControlTipText = GetText(g_arr_Text, "TIP050")
        .lbl_slipstream_2.ControlTipText = GetText(g_arr_Text, "TIP051")
        .chk_SlipstreamShow.ControlTipText = GetText(g_arr_Text, "TIP019")
        .chkRS_ro16.ControlTipText = GetText(g_arr_Text, "TIP020")
        .scrRS_ro03.ControlTipText = GetText(g_arr_Text, "TIP021a") & " " & g_arr_Grammar(4) & " " & GetText(g_arr_Text, "TIP021b")
        .cmdRS_ro03a.ControlTipText = GetText(g_arr_Text, "TIP022")
        .cmdRS_ro03b.ControlTipText = GetText(g_arr_Text, "TIP023")
        .chk_MOM1.ControlTipText = GetText(g_arr_Text, "TIP028")
        .chk_MOM2.ControlTipText = GetText(g_arr_Text, "TIP052")
        .chkRSMon.ControlTipText = GetText(g_arr_Text, "TIP053") & " " & g_arr_Grammar(8) & " " & GetText(g_arr_Text, "TIP054")
        .optRSMon1.ControlTipText = GetText(g_arr_Text, "TIP060")
        .optRSMon2.ControlTipText = GetText(g_arr_Text, "TIP061")
        .chkTribunes.ControlTipText = GetText(g_arr_Text, "TIP035")
        .chk_TEO1.ControlTipText = GetText(g_arr_Text, "TIP029")
        .chk_TEO2.ControlTipText = GetText(g_arr_Text, "TIP030")
        .chkRS_ro19.ControlTipText = GetText(g_arr_Text, "TIP036")
        .chkAutoSaveRace.ControlTipText = GetText(g_arr_Text, "TIP056") & " " & g_defaultAutoSavePath
        .opt_grid1.ControlTipText = GetText(g_arr_Text, "TIP001a") & " " & g_arr_Grammar(3) & " " & GetText(g_arr_Text, "RACEOPT098")
        .opt_grid2.ControlTipText = GetText(g_arr_Text, "TIP001a") & " " & g_arr_Grammar(3) & " " & GetText(g_arr_Text, "RACEOPT099")
        
        'Get values
        .optRS_ro01.Value = (objOption.TACTICS = False)
        .optRS_ro02.Value = (objOption.TACTICS = True)
        .chk_TEO1.Value = objOption.TACTICS_REVEAL_TAC
        .chk_TEO2.Value = objOption.TACTICS_REVEAL_CURR
        .chk_TEO1.Enabled = (objOption.TACTICS = True)  'Enabled only if the "Individual tactics" option is chosen
        .chk_TEO2.Enabled = (objOption.TACTICS = True)  'Enabled only if the "Individual tactics" option is chosen
        .opt_foc01.Value = objOption.FOCUSED_RUN = enumCamera.standard
        .opt_foc02.Value = objOption.FOCUSED_RUN = enumCamera.focus_horse
        .opt_foc03.Value = objOption.FOCUSED_RUN = enumCamera.focus_leader
        .chk_foc02b.Value = objOption.HIGHLIGHT_FOC
        .chk_foc02b.Enabled = (objOption.FOCUSED_RUN = enumCamera.focus_horse) 'Enabled only if the "Focused Run (Horse)" option is selected
        .chkRS_ro02.Value = objOption.HOOFPRINTS
        .scrRS_ro02.Value = objOption.METRES_DISPLAY
        .chkRS_ro04.Value = objOption.NAMES_LEFT
        .chkRS_ro05.Value = objOption.COLOURS_LEFT
        .chkRS_ro06.Value = objOption.HIGHLIGHT_FAV
        .chkFavAnn.Value = objOption.ANNOUNCE_FAV
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
        .scrSlipstream.Value = objOption.SLIPSTREAM_IMPACT
        .chk_SlipstreamShow.Value = objOption.SLIPSTREAM_SHOW
        .chk_SlipstreamShow.Enabled = (objOption.SLIPSTREAM_IMPACT > 0) 'Enabled only if slipstream impact is activated
        .scrRS_ro03.Value = objOption.SPEED_FACTOR
        .chkRS_ro15.Value = objOption.REFUSE_RUN
        .scrRefusalRate.Value = objOption.REFUSAL_RATE
        .scrRefusalRate.Enabled = (objOption.REFUSE_RUN = True) 'Enabled only if the "Refuse to run" checkbox is ticked
        .lblRefusalRate.Enabled = (objOption.REFUSE_RUN = True) 'Enabled only if the "Refuse to run" checkbox is ticked
        .chkRS_ro16.Value = objOption.SPEECH
        .chkTribunes.Value = objOption.TRIBUNES
        .scrSpec.Value = objOption.SPECTATORS 'in %
        .chk_MOM1.Value = objOption.MOMENTUM_BARS
        .chk_MOM2.Value = objOption.MOMENTUM_ICONS
        .scrMomRefr.Value = objOption.MOMENTUM_REFRESHRATE
        .scrMomRefr.Enabled = (objOption.MOMENTUM_BARS = True Or objOption.MOMENTUM_ICONS = True) 'Enabled only if one of the "Momentum" checkboxes is ticked
        .lblMomRefr.Enabled = (objOption.MOMENTUM_BARS = True Or objOption.MOMENTUM_ICONS = True) 'Enabled only if one of the "Momentum" checkboxes is ticked
        .lblRefrLow1.Enabled = (objOption.MOMENTUM_BARS = True Or objOption.MOMENTUM_ICONS = True) 'Enabled only if one of the "Momentum" checkboxes is ticked
        .lblRefrHigh1.Enabled = (objOption.MOMENTUM_BARS = True Or objOption.MOMENTUM_ICONS = True) 'Enabled only if one of the "Momentum" checkboxes is ticked
        .chkRSMon.Value = objOption.SPEEDMONITOR
        .scrRSMonRefr.Value = objOption.SPEEDMON_REFRESHRATE
        .scrRSMonRefr.Enabled = (objOption.SPEEDMONITOR = True) 'Enabled only if the "RSMon" checkbox is ticked
        .lblRSMonRefr.Enabled = (objOption.SPEEDMONITOR = True) 'Enabled only if the "RSMon" checkbox is ticked
        .lblRefrLow2.Enabled = (objOption.SPEEDMONITOR = True) 'Enabled only if the "RSMon" checkbox is ticked
        .lblRefrHigh2.Enabled = (objOption.SPEEDMONITOR = True) 'Enabled only if the "RSMon" checkbox is ticked
        .optRSMon1.Value = objOption.RSMON_SPEED
        .optRSMon2.Value = objOption.RSMON_DISTANCE
        .optRSMon1.Enabled = (objOption.SPEEDMONITOR = True) 'Enabled only if the "RSMon" checkbox is ticked
        .optRSMon2.Enabled = (objOption.SPEEDMONITOR = True) 'Enabled only if the "RSMon" checkbox is ticked
        .opt_grid1.Value = objOption.STARTING_GRID_IN
        .opt_grid2.Value = objOption.STARTING_GRID_BEHIND
        .chkRS_ro19.Value = objOption.AUTOFIT
        .chkAutoSaveRace.Value = objOption.AUTO_SAVE
    End With
    
    'Set the initial race info preview colours
    Call DefaultPreviewColours
    
    'Display the UserForm in the center of the Window
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
    objOption.TACTICS_REVEAL_TAC = chk_TEO1.Value
    objOption.TACTICS_REVEAL_CURR = chk_TEO2.Value
    objOption.SLIPSTREAM_IMPACT = scrSlipstream.Value
    objOption.SLIPSTREAM_SHOW = chk_SlipstreamShow.Value
    Select Case True
        Case opt_foc01.Value
            objOption.FOCUSED_RUN = enumCamera.standard
        Case opt_foc02.Value
            objOption.FOCUSED_RUN = enumCamera.focus_horse
        Case opt_foc03.Value
            objOption.FOCUSED_RUN = enumCamera.focus_leader
    End Select
    objOption.HIGHLIGHT_FOC = chk_foc02b.Value
    objOption.HOOFPRINTS = chkRS_ro02.Value
    objOption.METRES_DISPLAY = scrRS_ro02.Value
    objOption.NAMES_LEFT = chkRS_ro04.Value
    objOption.COLOURS_LEFT = chkRS_ro05.Value
    objOption.HIGHLIGHT_FAV = chkRS_ro06.Value
    objOption.ANNOUNCE_FAV = chkFavAnn.Value
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
    objOption.REFUSAL_RATE = scrRefusalRate.Value
    objOption.SPEECH = chkRS_ro16.Value
    objOption.MOMENTUM_BARS = chk_MOM1.Value
    objOption.MOMENTUM_ICONS = chk_MOM2.Value
    objOption.MOMENTUM_REFRESHRATE = scrMomRefr.Value
    objOption.SPEEDMONITOR = chkRSMon.Value
    objOption.RSMON_SPEED = optRSMon1
    objOption.RSMON_DISTANCE = optRSMon2
    objOption.SPEEDMON_REFRESHRATE = scrRSMonRefr.Value
    objOption.TRIBUNES = chkTribunes.Value
    objOption.SPECTATORS = scrSpec.Value
    objOption.AUTOFIT = chkRS_ro19.Value
    objOption.AUTO_SAVE = chkAutoSaveRace.Value
    objOption.STARTING_GRID_IN = opt_grid1.Value
    objOption.STARTING_GRID_BEHIND = opt_grid2.Value
    
    'Adapt the caption of the race start button ("Start the race" or "Betting and race")
    If g_strPlayMode = "RS" Then
        g_wksRace.OLEObjects("startrace").Object.caption = GetText(g_arr_Text, getCaptionStartBtn(objOption.BET_MODE))
    Else 'AI edition - refresh the ribbon
        g_RibbonGaloppSim.Invalidate
    End If
    
    'Close the UserForm
    Unload Me
End Sub

'Click on the "Focused Run (Standard)" radio button
Private Sub opt_foc01_Click()
    Me.chk_foc02b.Enabled = False
End Sub

'Click on the "Focused Run (Horse)" radio button
Private Sub opt_foc02_Click()
    Me.chk_foc02b.Enabled = True
End Sub

'Click on the "Focused Run (Leader)" radio button
Private Sub opt_foc03_Click()
    Me.chk_foc02b.Enabled = False
End Sub

'Click on the "Placing bets" checkbox
Private Sub chkRS_ro11a_Click()
    Me.chkRS_ro11b.Enabled = (Me.chkRS_ro11a.Value = True) 'Set the status dependent on the "Placing bets" main checkbox
End Sub

'Click on the "Momentum speed bars" checkbox
Private Sub chk_MOM1_Click()
    With Me 'Set the status dependent on the checkbox above
        .scrMomRefr.Enabled = (Me.chk_MOM1.Value = True Or Me.chk_MOM2.Value = True)
        .lblMomRefr.Enabled = (Me.chk_MOM1.Value = True Or Me.chk_MOM2.Value = True)
        .lblRefrLow1.Enabled = (Me.chk_MOM1.Value = True Or Me.chk_MOM2.Value = True)
        .lblRefrHigh1.Enabled = (Me.chk_MOM1.Value = True Or Me.chk_MOM2.Value = True)
    End With
End Sub

'Click on the "Momentum icons" checkbox
Private Sub chk_MOM2_Click()
    With Me 'Set the status dependent on the checkbox above
        .scrMomRefr.Enabled = (Me.chk_MOM2.Value = True Or Me.chk_MOM1.Value = True)
        .lblMomRefr.Enabled = (Me.chk_MOM2.Value = True Or Me.chk_MOM1.Value = True)
        .lblRefrLow1.Enabled = (Me.chk_MOM2.Value = True Or Me.chk_MOM1.Value = True)
        .lblRefrHigh1.Enabled = (Me.chk_MOM2.Value = True Or Me.chk_MOM1.Value = True)
    End With
End Sub

'Click on the "RSMon" checkbox
Private Sub chkRSMon_Click()
   With Me 'Set the status dependent on the checkbox above
    .scrRSMonRefr.Enabled = (.chkRSMon.Value = True)
    .lblRSMonRefr.Enabled = (.chkRSMon.Value = True)
    .lblRefrLow2.Enabled = (.chkRSMon.Value = True)
    .lblRefrHigh2.Enabled = (.chkRSMon.Value = True)
    .optRSMon1.Enabled = (.chkRSMon.Value = True)
    .optRSMon2.Enabled = (.chkRSMon.Value = True)
End With
End Sub

'Change of the slipstream slider position
Private Sub scrSlipstream_Change()
    With Me
        .chk_SlipstreamShow.Enabled = (.scrSlipstream.Value > 0) 'Activate dependent on the slipstream slider
    End With
End Sub

'Click on the "Refuse to run" checkbox
Private Sub chkRS_ro15_Click()
    With Me 'Set the status dependent on the checkbox above
        .scrRefusalRate.Enabled = (Me.chkRS_ro15.Value = True)
        .lblRefusalRate.Enabled = (Me.chkRS_ro15.Value = True)
    End With
End Sub

'Change of the refusal rate slider position
Private Sub scrRefusalRate_Change()
    'Adapt the text
    lblRefusalRate.caption = "1 " & GetText(g_arr_Text, "RACEOPT065") & " " & scrRefusalRate.Value & " " & GetText(g_arr_Text, "RACEOPT066")
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

'Click on the "No tactics" radio button
Private Sub optRS_ro01_Click()
    'Diasble the tactics revelation checkboxes
    chk_TEO1.Enabled = False
    chk_TEO2.Enabled = False
End Sub

'Click on the "Individual tactics" radio button
Private Sub optRS_ro02_Click()
    'Enasble the tactics revelation checkboxes
    chk_TEO1.Enabled = True
    chk_TEO2.Enabled = True
End Sub

'Click on the "Save settings" button
Private Sub cmdRS_ro03a_Click()
    Call basInputOutput.SaveRaceOptions
End Sub

'Click on the "Load settings" button
Private Sub cmdRS_ro03b_Click()
    Call basInputOutput.LoadRaceOptions
End Sub
