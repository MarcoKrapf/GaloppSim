VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionsRace 
   Caption         =   "[Race options]"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9420
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
    
    'Settings of the zoom level slider
        With scrRS_ro01
            .min = 1 'minumum value
            .max = 3 'maxmum value
            .SmallChange = 1 'value change when using the arrows
            .LargeChange = 1 'value change when clicking inside the slider
        End With
    
    With Me
        'Captions
        .Caption = g_strTxt(252)
        .lblRS_ro01.Caption = g_strTxt(292)
        .lblRS_ro02.Caption = g_strTxt(293)
        .lblRS_ro03.Caption = g_strTxt(294)
        .lblRS_ro05.Caption = g_strTxt(104)
        .fraRS_ro01.Caption = g_strTxt(260)
        .optRS_ro01.Caption = g_strTxt(261)
        .optRS_ro02.Caption = g_strTxt(262)
        .optRS_ro03.Caption = g_strTxt(263)
        .chkRS_ro01a.Caption = g_strTxt(250)
        .chkRS_ro01b.Caption = g_strTxt(255)
        .chkRS_ro02.Caption = g_strTxt(248)
        .chkRS_ro04.Caption = g_strTxt(256)
        .chkRS_ro05.Caption = g_strTxt(257)
        .chkRS_ro06.Caption = g_strTxt(258)
        .chkRS_ro07.Caption = g_strTxt(259)
        .chkRS_ro08.Caption = g_strTxt(266)
        .chkRS_ro09.Caption = g_strTxt(267)
        .chkRS_ro10.Caption = g_strTxt(264)
        .chkRS_ro11a.Caption = g_strTxt(450)
        .chkRS_ro11b.Caption = g_strTxt(452)
        .cmdRS_ro01.Caption = g_strTxt(180)
        .lblRS_ro04.Caption = g_strTxt(369)
        .fraRS_ro02.Caption = g_strTxt(370) & ": " & ZoomLevelText(g_byteTrackZoom)
        .lblscrRS_ro02.Caption = g_strTxt(374)
        'ControlTipTexts
        .optRS_ro01.ControlTipText = g_strTxt(299)
        .optRS_ro02.ControlTipText = g_strTxt(301)
        .optRS_ro03.ControlTipText = g_strTxt(301)
        .chkRS_ro01a.ControlTipText = g_strTxt(343)
        .chkRS_ro01b.ControlTipText = "xxxxxxxxxxxxxxxxxxx" 'ToDo
        .chkRS_ro02.ControlTipText = g_strTxt(303)
        .chkRS_ro04.ControlTipText = g_strTxt(307)
        .chkRS_ro05.ControlTipText = g_strTxt(309)
        .chkRS_ro06.ControlTipText = g_strTxt(311)
        .chkRS_ro07.ControlTipText = g_strTxt(313)
        .chkRS_ro08.ControlTipText = g_strTxt(315)
        .chkRS_ro09.ControlTipText = g_strTxt(317)
        .chkRS_ro10.ControlTipText = g_strTxt(357)
        .chkRS_ro11a.ControlTipText = g_strTxt(451)
        .chkRS_ro11b.ControlTipText = g_strTxt(453)
        'Get values
        .optRS_ro01.Value = g_blnTactics0
        .optRS_ro02.Value = g_blnTactics3
        .optRS_ro03.Value = g_blnTactics6
        .chkRS_ro01a.Value = g_blnFocusedRun
        .chkRS_ro01b.Value = g_blnHighlightFoc
        .chkRS_ro01b.Enabled = (g_blnFocusedRun = True) 'enabled only if the "Focused Run" checkbox is ticked
        .chkRS_ro02.Value = g_blnHoofprints
        .scrRS_ro01.Value = g_byteTrackZoom
        .chkRS_ro04.Value = g_blnHorseNamesLeft
        .chkRS_ro05.Value = g_blnHorseColoursLeft
        .chkRS_ro06.Value = g_blnHighlightFav
        .chkRS_ro07.Value = g_blnHorseNamesFinish
        .chkRS_ro08.Value = g_blnRankingColours
        .chkRS_ro09.Value = g_blnRankingDelay
        .chkRS_ro10.Value = g_blnRaceInformation
        .chkRS_ro11a.Value = g_blnBettingMode
        .chkRS_ro11b.Value = g_blnBettingAnalysis
        .chkRS_ro11b.Enabled = (g_blnBettingMode = True) 'enabled only if the "Placing bets" checkbox is ticked
    End With
    
    'Display the UserForm in the center of the Window
    Call basMainCode.PlaceUserFormInCenter(Me)
End Sub

'OK button
Private Sub cmdRS_ro01_Click()
    'Set values
    g_blnTactics0 = optRS_ro01.Value
    g_blnTactics3 = optRS_ro02.Value
    g_blnTactics6 = optRS_ro03.Value
    g_blnFocusedRun = chkRS_ro01a.Value
    g_blnHighlightFoc = chkRS_ro01b.Value
    g_blnHoofprints = chkRS_ro02.Value
    g_byteTrackZoom = scrRS_ro01.Value
    g_blnHorseNamesLeft = chkRS_ro04.Value
    g_blnHorseColoursLeft = chkRS_ro05.Value
    g_blnHighlightFav = chkRS_ro06.Value
    g_blnHorseNamesFinish = chkRS_ro07.Value
    g_blnRankingColours = chkRS_ro08.Value
    g_blnRankingDelay = chkRS_ro09.Value
    g_blnRaceInformation = chkRS_ro10.Value
    g_blnBettingMode = chkRS_ro11a.Value
    g_blnBettingAnalysis = chkRS_ro11b.Value
    'Close UserForm
    Unload Me
End Sub

'Click on the "Focused Run" checkbox
Private Sub chkRS_ro01a_Click()
    Me.chkRS_ro01b.Enabled = (Me.chkRS_ro01a.Value = True) 'set the status dependent on the "Focused Run" checkbox
End Sub

'Click on the "Placing bets" checkbox
Private Sub chkRS_ro11a_Click()
    Me.chkRS_ro11b.Enabled = (Me.chkRS_ro11a.Value = True) 'set the status dependent on the "Placing bets" checkbox
End Sub

'Change of the zoom level slider position
Private Sub scrRS_ro01_Change()
    'Adapt the horse size preview
    lblscrRS_ro01.Width = basMainCode.HorseSizePreview(scrRS_ro01.Value)(0)
    lblscrRS_ro01.Height = basMainCode.HorseSizePreview(scrRS_ro01.Value)(1)
    'Adapt the zoom level text
        fraRS_ro02.Caption = g_strTxt(370) & ": " & basMainCode.ZoomLevelText(scrRS_ro01.Value)
End Sub
