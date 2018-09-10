VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionsRace 
   Caption         =   "[Race options]"
   ClientHeight    =   8520
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

Private Sub UserForm_Initialize()

    'Get the zoom level recommendation
        g_byteTrackZoom = basMainCode.ZoomLevelRecommendation
    
    'Settings of the sliders
        'Zoom level slider
        With scrRS_ro01
            .min = 1 'minumum value
            .max = 3 'maxmum value
            .SmallChange = 1 'value change when using the arrows
            .LargeChange = 1 'value change when clicking inside the slider
        End With
        'Metres
        With scrRS_ro02
            .min = 50 'minumum value
            .max = 1000 'maxmum value
            .SmallChange = 50 'value change when using the arrows
            .LargeChange = 50 'value change when clicking inside the slider
        End With
        'Race speed factor
        With scrRS_ro03
            .min = 1 'minumum value
            .max = 5 'maxmum value
            .SmallChange = 1 'value change when using the arrows
            .LargeChange = 1 'value change when clicking inside the slider
        End With
        
        Call raceSpeedFactors
    
    With Me
        'Captions
        .Caption = GetTxt(g_arrTxt, "USERFORM001")
        .lblRS_ro01.Caption = GetTxt(g_arrTxt, "RACEOPT001")
        .lblRS_ro05.Caption = GetTxt(g_arrTxt, "START005")
        .fraRS_ro01.Caption = GetTxt(g_arrTxt, "RACEOPT009")
        .optRS_ro01.Caption = GetTxt(g_arrTxt, "RACEOPT010")
        .optRS_ro02.Caption = GetTxt(g_arrTxt, "RACEOPT011")
        .optRS_ro03.Caption = GetTxt(g_arrTxt, "RACEOPT012")
        .chkRS_ro02.Caption = GetTxt(g_arrTxt, "RACEOPT017")
        .chkRS_ro08.Caption = GetTxt(g_arrTxt, "RACEOPT015")
        .chkRS_ro09.Caption = GetTxt(g_arrTxt, "RACEOPT016")
        .lblRS_ro06.Caption = GetTxt(g_arrTxt, "RACEOPT013")
        .fraRS_ro05.Caption = GetTxt(g_arrTxt, "RACEOPT013")
        .chkRS_ro10.Caption = GetTxt(g_arrTxt, "RACEOPT014")
        .chkRS_ro13.Caption = GetTxt(g_arrTxt, "RACEOPT032")
        .chkRS_ro14.Caption = GetTxt(g_arrTxt, "RACEOPT034")
        .chkRS_ro11a.Caption = GetTxt(g_arrTxt, "RACEOPT022")
        .chkRS_ro11b.Caption = GetTxt(g_arrTxt, "RACEOPT024")
        .fraRS_ro03.Caption = GetTxt(g_arrTxt, "RACEOPT018")
        .chkRS_ro12a.Caption = GetTxt(g_arrTxt, "RACEOPT019")
        .chkRS_ro12b.Caption = GetTxt(g_arrTxt, "RACEOPT020")
        .chkRS_ro12c.Caption = GetTxt(g_arrTxt, "RACEOPT021")
        .fraRS_ro04.Caption = GetTxt(g_arrTxt, "RACEOPT026")
        .chkRS_ro01a.Caption = GetTxt(g_arrTxt, "RACEOPT027")
        .chkRS_ro01b.Caption = GetTxt(g_arrTxt, "RACEOPT004")
        .cmdRS_ro01.Caption = GetTxt(g_arrTxt, "BTN014")
        .lblRS_ro08.Caption = GetTxt(g_arrTxt, "RACEOPT035")
        .fraRS_ro02.Caption = GetTxt(g_arrTxt, "ZOOM001") & ": " & ZoomLevelText(g_byteTrackZoom)
        .lblscrRS_ro02.Caption = GetTxt(g_arrTxt, "ZOOM006")
        .fraRS_ro07.Caption = GetTxt(g_arrTxt, "RACEOPT036") & " " _
                                & GetTxt(g_arrTxt, "RACE020") & " " & scrRS_ro02.Value & GetTxt(g_arrTxt, "RACE008")
        .cmdRS_ro02a.Caption = GetTxt(g_arrTxt, "RACEOPT029")
        .cmdRS_ro02b.Caption = GetTxt(g_arrTxt, "RACEOPT030")
        .cmdRS_ro02c.Caption = GetTxt(g_arrTxt, "RACEOPT031")
        .lblRS_ro07.Caption = GetTxt(g_arrTxt, "RACEOPT033")
        .fraRS_ro08.Caption = GetTxt(g_arrTxt, "RACEOPT002")
        .chkRS_ro04.Caption = GetTxt(g_arrTxt, "RACEOPT005")
        .chkRS_ro05.Caption = GetTxt(g_arrTxt, "RACEOPT006")
        .chkRS_ro06.Caption = GetTxt(g_arrTxt, "RACEOPT007")
        .chkRS_ro07.Caption = GetTxt(g_arrTxt, "RACEOPT008")
        .fraRS_ro09.Caption = GetTxt(g_arrTxt, "RACEOPT003")
        .lblRS_ro09.Caption = GetTxt(g_arrTxt, "RACEOPT041")
        .cmdRS_ro03a.Caption = GetTxt(g_arrTxt, "RACEOPT037")
        .cmdRS_ro03b.Caption = GetTxt(g_arrTxt, "RACEOPT038")
        .opt_ro01.Caption = GetTxt(g_arrTxt, "RACEOPT039")
        .opt_ro02.Caption = GetTxt(g_arrTxt, "RACEOPT040")
        .lblRS_ro10.Caption = GetTxt(g_arrTxt, "RACEOPT042")
        .lblRS_ro11.Caption = g_intGMPspeedfactor & GetTxt(g_arrTxt, "RACEOPT043")
        .lblRS_ro12a.Caption = GetTxt(g_arrTxt, "RACEOPT046")
        .lblRS_ro12b.Caption = GetTxt(g_arrTxt, "RACEOPT047")
        .fraRS_ro12.Caption = GetTxt(g_arrTxt, "RACEOPT044")
        .chkRS_ro15.Caption = GetTxt(g_arrTxt, "RACEOPT045")
        .chkRS_ro16.Caption = GetTxt(g_arrTxt, "RACEOPT048")
        
        'ControlTipTexts
        .optRS_ro01.ControlTipText = GetTxt(g_arrTxt, "TIP001")
        .optRS_ro02.ControlTipText = GetTxt(g_arrTxt, "TIP002")
        .optRS_ro03.ControlTipText = GetTxt(g_arrTxt, "TIP002")
        .chkRS_ro01a.ControlTipText = GetTxt(g_arrTxt, "TIP003")
        .chkRS_ro02.ControlTipText = GetTxt(g_arrTxt, "TIP004")
        .chkRS_ro04.ControlTipText = GetTxt(g_arrTxt, "TIP005")
        .chkRS_ro05.ControlTipText = GetTxt(g_arrTxt, "TIP006")
        .chkRS_ro06.ControlTipText = GetTxt(g_arrTxt, "TIP007")
        .chkRS_ro07.ControlTipText = GetTxt(g_arrTxt, "TIP008")
        .chkRS_ro08.ControlTipText = GetTxt(g_arrTxt, "TIP009")
        .chkRS_ro09.ControlTipText = GetTxt(g_arrTxt, "TIP010")
        .chkRS_ro10.ControlTipText = GetTxt(g_arrTxt, "TIP011")
        .chkRS_ro11a.ControlTipText = GetTxt(g_arrTxt, "RACEOPT023")
        .chkRS_ro11b.ControlTipText = GetTxt(g_arrTxt, "RACEOPT025")
        .chkRS_ro15.ControlTipText = GetTxt(g_arrTxt, "TIP016")
        .chkRS_ro12a.ControlTipText = GetTxt(g_arrTxt, "TIP017")
        .chkRS_ro12b.ControlTipText = GetTxt(g_arrTxt, "TIP018")
        .chkRS_ro12c.ControlTipText = GetTxt(g_arrTxt, "TIP019")
        .chkRS_ro16.ControlTipText = GetTxt(g_arrTxt, "TIP020")
        .scrRS_ro03.ControlTipText = GetTxt(g_arrTxt, "TIP021")
        .cmdRS_ro03a.ControlTipText = GetTxt(g_arrTxt, "TIP022")
        .cmdRS_ro03b.ControlTipText = GetTxt(g_arrTxt, "TIP023")
        
        'Get values
        .optRS_ro01.Value = g_blnTactics0
        .optRS_ro02.Value = g_blnTactics3
        .optRS_ro03.Value = g_blnTactics6
        .chkRS_ro01a.Value = g_blnFocusedRun
        .chkRS_ro01b.Value = g_blnHighlightFoc
        .chkRS_ro01b.Enabled = (g_blnFocusedRun = True) 'enabled only if the "Focused Run" checkbox is ticked
        .chkRS_ro02.Value = g_blnHoofprints
        .scrRS_ro01.Value = g_byteTrackZoom
        .scrRS_ro02.Value = g_intTrackMetres
        .chkRS_ro04.Value = g_blnHorseNamesLeft
        .chkRS_ro05.Value = g_blnHorseColoursLeft
        .chkRS_ro06.Value = g_blnHighlightFav
        .chkRS_ro07.Value = g_blnHorseNamesFinish
        .chkRS_ro08.Value = g_blnRankingColours
        .chkRS_ro09.Value = g_blnRankingDelay
        .chkRS_ro10.Value = g_blnRaceInformation
        .opt_ro01.Value = g_blnRaceInfoPopup
        .opt_ro02.Value = g_blnRaceInfoWorksheet
        .opt_ro01.Enabled = (g_blnRaceInformation = True)  'enabled only if the "Race info" checkbox is ticked
        .opt_ro02.Enabled = (g_blnRaceInformation = True)  'enabled only if the "Race info" checkbox is ticked
        .chkRS_ro13.Value = g_blnRaceInfoLeader
        .chkRS_ro13.Enabled = (g_blnRaceInformation = True)  'enabled only if the "Race info" checkbox is ticked
        .chkRS_ro14.Value = g_blnRaceInfoProgressBar
        .chkRS_ro14.Enabled = (g_blnRaceInformation = True)  'enabled only if the "Race info" checkbox is ticked
        cmdRS_ro02a.Enabled = (g_blnRaceInformation = True)  'enabled only if the "Race info" checkbox is ticked
        cmdRS_ro02b.Enabled = (g_blnRaceInformation = True)  'enabled only if the "Race info" checkbox is ticked
        cmdRS_ro02c.Enabled = (g_blnRaceInformation = True)  'enabled only if the "Race info" checkbox is ticked
        lblRS_ro07.BackColor = g_lngRaceInfoBackColour
        lblRS_ro07.ForeColor = g_lngRaceInfoForeColour
        lblRS_ro07.Visible = (g_blnRaceInformation = True)  'enabled only if the "Race info" checkbox is ticked
        .chkRS_ro11a.Value = g_blnBettingMode
        .chkRS_ro11b.Value = g_blnBettingAnalysis
        .chkRS_ro11b.Enabled = (g_blnBettingMode = True) 'enabled only if the "Placing bets" checkbox is ticked
        .chkRS_ro12a.Value = g_blnSlipstream
        .chkRS_ro12b.Value = g_blnSlipstreamDouble
        .chkRS_ro12b.Enabled = (g_blnSlipstream = True) 'enabled only if the "Slipstreaming" checkbox is ticked
        .chkRS_ro12c.Value = g_blnSlipstreamShow
        .chkRS_ro12c.Enabled = (g_blnSlipstream = True) 'enabled only if the "Slipstreaming" checkbox is ticked
        .scrRS_ro03.Value = g_intGMPspeedfactor
        .chkRS_ro15.Value = g_blnIncidentRefuse
        .chkRS_ro16.Value = g_blnSpeech
    End With
    
    'Set the initial race info preview colours
    Call DefaultPreviewColours
    
    'Display the UserForm in the center of the Window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

'Option button "Race info in a pop-up"
Private Sub opt_ro01_Click()
    Call DefaultPreviewColours
    'Re-calculate the overall race speed
        Call raceSpeedFactors
End Sub

'Option button "Race info on the worksheet"
Private Sub opt_ro02_Click()
    Call DefaultPreviewColours
    'Re-calculate the overall race speed
        Call raceSpeedFactors
End Sub

'Click on the "Leader" checkbox
Private Sub chkRS_ro13_Click()
    'Re-calculate the overall race speed
        Call raceSpeedFactors
End Sub

'Click on the "Race progress" checkbox
Private Sub chkRS_ro14_Click()
    'Re-calculate the overall race speed
        Call raceSpeedFactors
End Sub

'Default colours of the race info preview
Private Sub DefaultPreviewColours()
    If opt_ro01 Then 'in a pop-up
        If lblRS_ro07.BackColor = vbWhite Then lblRS_ro07.BackColor = -2147483633
        If lblRS_ro07.ForeColor = vbBlack Then lblRS_ro07.ForeColor = -2147483630
    End If
    If opt_ro02 Then 'on the worksheet
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
    If opt_ro01 Then 'in a pop-up
        lblRS_ro07.BackColor = cmdRS_ro02c.BackColor
        lblRS_ro07.ForeColor = cmdRS_ro02c.ForeColor
    End If
    If opt_ro02 Then 'on the worksheet
        lblRS_ro07.BackColor = vbWhite
        lblRS_ro07.ForeColor = vbBlack
    End If
End Sub

'OK button
Private Sub cmdRS_ro01_Click()
    'Set values
    g_blnTactics0 = optRS_ro01.Value
    g_blnTactics3 = optRS_ro02.Value
    g_blnTactics6 = optRS_ro03.Value
    g_blnSlipstream = chkRS_ro12a.Value
    g_blnSlipstreamDouble = chkRS_ro12b.Value
    g_blnSlipstreamShow = chkRS_ro12c.Value
    g_blnFocusedRun = chkRS_ro01a.Value
    g_blnHighlightFoc = chkRS_ro01b.Value
    g_blnHoofprints = chkRS_ro02.Value
    g_byteTrackZoom = scrRS_ro01.Value
    g_intTrackMetres = scrRS_ro02.Value
    g_blnHorseNamesLeft = chkRS_ro04.Value
    g_blnHorseColoursLeft = chkRS_ro05.Value
    g_blnHighlightFav = chkRS_ro06.Value
    g_blnHorseNamesFinish = chkRS_ro07.Value
    g_blnRankingColours = chkRS_ro08.Value
    g_blnRankingDelay = chkRS_ro09.Value
    g_blnRaceInformation = chkRS_ro10.Value
    g_blnRaceInfoPopup = opt_ro01.Value
    g_blnRaceInfoWorksheet = opt_ro02.Value
    g_blnRaceInfoLeader = chkRS_ro13.Value
    g_blnRaceInfoProgressBar = chkRS_ro14.Value
    g_lngRaceInfoBackColour = lblRS_ro07.BackColor
    g_lngRaceInfoForeColour = lblRS_ro07.ForeColor
    g_blnBettingMode = chkRS_ro11a.Value
    g_blnBettingAnalysis = chkRS_ro11b.Value
    g_intGMPspeedfactor = scrRS_ro03.Value
    g_blnIncidentRefuse = chkRS_ro15.Value
    g_blnSpeech = chkRS_ro16.Value
    
    'Adapt the caption of the race start button
    If g_strPlayMode = "RS" Then
        g_wksRace.OLEObjects("startrace").Object.Caption = GetTxt(g_arrTxt, getCaptionStartBtn(g_blnBettingMode))
    Else 'AI edition - refresh the ribbon
        g_RibbonGaloppSim.Invalidate
    End If
    
    'Close UserForm
    Unload Me
End Sub

'Click on the "Hoof prints" checkbox
Private Sub chkRS_ro02_Click()
    'Re-calculate the overall race speed
        Call raceSpeedFactors
End Sub

'Click on the "Focused Run" checkbox
Private Sub chkRS_ro01a_Click()
    Me.chkRS_ro01b.Enabled = (Me.chkRS_ro01a.Value = True) 'set the status dependent on the "Focused Run" checkbox
    'Re-calculate the overall race speed
        Call raceSpeedFactors
End Sub

'Click on the "Placing bets" checkbox
Private Sub chkRS_ro11a_Click()
    Me.chkRS_ro11b.Enabled = (Me.chkRS_ro11a.Value = True) 'set the status dependent on the "Placing bets" checkbox
End Sub

'Click on the "Slipstreaming" checkbox
Private Sub chkRS_ro12a_Click()
    With Me
        .chkRS_ro12b.Enabled = (.chkRS_ro12a.Value = True) 'set the status dependent on the "Slipstreaming" checkbox
        .chkRS_ro12c.Enabled = (.chkRS_ro12a.Value = True) 'set the status dependent on the "Slipstreaming" checkbox
    End With
End Sub

'Click on the "Race info" checkbox
Private Sub chkRS_ro10_Click()
    'Set the status dependent on the "Race info" checkbox
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
    'Re-calculate the overall race speed
        Call raceSpeedFactors
End Sub

'Change of the zoom level slider position
Private Sub scrRS_ro01_Change()
    'Adapt the horse size preview
    lblscrRS_ro01.Width = basMainCode.HorseSizePreview(scrRS_ro01.Value)(0)
    lblscrRS_ro01.Height = basMainCode.HorseSizePreview(scrRS_ro01.Value)(1)
    'Adapt the zoom level text
        fraRS_ro02.Caption = GetTxt(g_arrTxt, "ZOOM001") & ": " & basMainCode.ZoomLevelText(scrRS_ro01.Value)
End Sub

'Change of the race track metres slider position
Private Sub scrRS_ro02_Change()
    Dim marker As String
    Dim i As Integer, j As Integer
    For i = 1 To (4000 / scrRS_ro02.Value)
        marker = marker & (i * scrRS_ro02.Value) & GetTxt(g_arrTxt, "RACE008")
        For j = 1 To (scrRS_ro02.Value / scrRS_ro02.min)
            marker = marker & " "
        Next
    Next
    'Adapt the text
    lblscrRS_ro03.Caption = marker
    fraRS_ro07.Caption = GetTxt(g_arrTxt, "RACEOPT036") & " " _
                        & GetTxt(g_arrTxt, "RACE020") & " " & scrRS_ro02.Value & GetTxt(g_arrTxt, "RACE008")
End Sub

'Change of the race speed factor slider position
Private Sub scrRS_ro03_Change()
    'Adapt the race speed factor label
        lblRS_ro11.Caption = scrRS_ro03.Value & GetTxt(g_arrTxt, "RACEOPT043")
    'Re-calculate the overall race speed
        Call raceSpeedFactors
End Sub

'Click on the "Save settings" button
Private Sub cmdRS_ro03a_Click()
    Call basInputOutput.SaveRaceOptions
End Sub

'Click on the "Load settings" button
Private Sub cmdRS_ro03b_Click()
    Call basInputOutput.LoadRaceOptions
End Sub

'Race speed factors
Private Sub raceSpeedFactors()
    'Race info factor
        If chkRS_ro10.Value = True Then 'race information is displayed during the race
            If opt_ro01.Value = True Then 'in a pop-up
                g_arr_RaceSpeed(0) = 0
                If chkRS_ro13.Value = True Then g_arr_RaceSpeed(0) = 0 'leader
                If chkRS_ro14.Value = True Then g_arr_RaceSpeed(0) = -1 'race progress
            End If
            If opt_ro02.Value = True Then 'on the worksheet
                g_arr_RaceSpeed(0) = 0
                If chkRS_ro13.Value = True Then g_arr_RaceSpeed(0) = -1 'leader
                If chkRS_ro14.Value = True Then g_arr_RaceSpeed(0) = -2 'race progress
            End If
        Else 'no race information during the race
            g_arr_RaceSpeed(0) = 0
        End If
    'Hoof print factor
        If chkRS_ro02.Value = True Then
            g_arr_RaceSpeed(1) = -2 * UBound(g_arr_RaceSpeed)
        Else
            g_arr_RaceSpeed(1) = 0
        End If
    'Focused run factor
        If chkRS_ro01a.Value = True Then
            g_arr_RaceSpeed(2) = -2 * UBound(g_arr_RaceSpeed)
        Else
            g_arr_RaceSpeed(2) = 0
        End If
    'GMSspeed factor
        g_arr_RaceSpeed(3) = (scrRS_ro03.Value - 1) * 3
    'Number of horses factor
g_arr_RaceSpeed(4) = 0
    
End Sub
