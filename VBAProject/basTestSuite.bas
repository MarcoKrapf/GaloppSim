Attribute VB_Name = "basTestSuite"
Option Explicit
Option Private Module

'This module contains non-productive procedures
'which can be used for automatic testing
'   Module basTestSuite

    'General parameters
    Dim tempSkipDelay As Boolean
    Dim tempLanguage As String
    Dim tempColourMode As String
    Dim tempDaylight As Integer
    
    'Race parameters
    Dim tempSPEED_FACTOR As Integer
    Dim tempSPEECH As Boolean
    Dim tempMOMENTUM As Boolean
    Dim tempMOMENTUM_REFRESHRATE As Integer
    Dim tempTACTICS As Boolean
    Dim tempTACTICS_REVEAL_TAC As Boolean
    Dim tempTACTICS_REVEAL_CURR As Boolean
    Dim tempREFUSE_RUN As Boolean
    Dim tempREFUSAL_RATE As Integer
    Dim tempSLIPSTREAM As Boolean
    Dim tempSLIPSTREAM_DBL As Boolean
    Dim tempSLIPSTREAM_SHOW As Boolean
    Dim tempFOCUSED_RUN As Long
    Dim tempHIGHLIGHT_FOC As Boolean
    Dim tempMETRES_DISPLAY As Integer
    Dim tempTRIBUNES As Boolean
    Dim tempSPECTATORS As Integer
    Dim tempHOOFPRINTS As Boolean
    Dim tempNAMES_LEFT As Boolean
    Dim tempCOLOURS_LEFT As Boolean
    Dim tempHIGHLIGHT_FAV As Boolean
    Dim tempNAMES_FINISH As Boolean
    Dim tempNAMES_PHOTO As Boolean
    Dim tempPHOTO_BW As Boolean
    Dim tempRANKING_COL As Boolean
    Dim tempRANKING_DELAY As Boolean
    Dim tempRACE_INFO As Boolean
    Dim tempRACE_INFO_POP As Boolean
    Dim tempRACE_INFO_WKS As Boolean
    Dim tempRACE_INFO_LEADER As Boolean
    Dim tempRACE_INFO_PROGRESS As Boolean
    Dim tempRACE_INFO_COL_B As Long
    Dim tempRACE_INFO_COL_F As Long
    
    'Race specific parameters
    Dim tempPARTICULATES_SLIDER As Integer
    Dim tempTIDE As Integer
    Dim tempLUGWORMS As Integer

'Execute this procedure for automatic testing
Public Sub TestSuite()
    frmTestSuite.show (vbModeless)
End Sub

Public Sub TestStart_std(colScope As Collection, blnRnd As Boolean, _
        blnCurrentSpeed As Boolean, blnNoMomentum As Boolean, blnNoSpeech As Boolean, _
        blnNoInfoCol As Boolean, blnCurrentColourMode As Boolean, _
        blnCurrentLanguage As Boolean)

    Dim TestCase As Integer

    Call RememberRaceOptions
    
    'Checkbox "Skip delay commands"
    If frmTestSuite.chkSkipDelay Then g_skipDelay = True Else g_skipDelay = False
    
    Debug.Print vbNewLine & vbNewLine & ">> AUTOMATIC TESTING - STANDARD" & vbNewLine

    'Play all races one by one
    For TestCase = 1 To colScope.count 'Loop through the test scope
        Set g_wksRaceData = ThisWorkbook.Worksheets(colScope(TestCase)) 'Get the next race
        objOption.ZOOM_LEVEL = ZoomLevelRecommendation(True) 'Get the recommended zoom value
        
        Debug.Print vbTab & "> Test " & TestCase & "/" & colScope.count & " - " & g_wksRaceData.name
        
        Call basAuxiliary.Scroll(1, 1) 'Scroll to the upper left

        If blnRnd Then
            Call TestAutomationRandomSettings
            If blnCurrentSpeed Then objOption.SPEED_FACTOR = tempSPEED_FACTOR
            If blnNoSpeech Then objOption.SPEECH = False
            If blnNoMomentum Then objOption.MOMENTUM = False
            If blnNoInfoCol Then
                objOption.RACE_INFO_COL_F = vbBlack
                If g_strPlayMode = "RS" Then
                    objOption.RACE_INFO_COL_B = vbWhite
                Else
                    objOption.RACE_INFO_COL_B = &H8000000F
                End If
            End If
            If blnCurrentLanguage Then
                objOption.language = tempLanguage
                Call GetTextComponents
                Call GetAnimalGrammar
            End If
            If blnCurrentColourMode Then
                g_strColourMode = tempColourMode
                'In case of 24h colour mode:
                    'Comment in for preserving the selected daylight
                    'Comment out for random daylight
'                objOption.DAYLIGHT = tempDaylight
            End If
        End If
        
        frmTestSuite.frameTestProgress.caption = "TEST RACE " & TestCase & "/" & colScope.count
        With frmTestSuite.lblTestProgress
            .caption = "RUNNING"
            .width = 348 / colScope.count * TestCase
            .BackColor = vbBlue
            .ForeColor = vbWhite
        End With

        Call ShowNewRaceScreen("RTA2") 'RTA title screen
        Application.Wait (Now + TimeValue("0:00:02"))
        Call NewRace(True) 'Start the test race
        Application.Wait (Now + TimeValue("0:00:02"))
        Call ShowRankingList(True)
        Application.Wait (Now + TimeValue("0:00:02"))
        Call ShowWinnerPhoto(True)
        Application.Wait (Now + TimeValue("0:00:02"))
        Call ShowFinishPhoto(True)
        Application.Wait (Now + TimeValue("0:00:02"))
        Call basAuxiliary.Freeze(0, 0, False) 'Unfreeze
        
        Debug.Print vbTab & "Test " & TestCase & "/" & colScope.count & " finished <" & vbNewLine

    Next
    
    Call basAuxiliary.Scroll(1, 1) 'Scroll to the upper left
    Call RS_MenuAreaShow(False)
    Call RestoreRaceOptions
    
    frmTestSuite.frameTestProgress.caption = ""
    With frmTestSuite.lblTestProgress
        .caption = "FINISHED"
        .BackColor = vbGreen
        .ForeColor = vbBlack
    End With
    
    Debug.Print "AUTOMATIC TESTING FINISHED <<" & vbNewLine & vbNewLine

End Sub

'Generate random settings for a test race
Private Sub TestAutomationRandomSettings()

    Dim languageRnd As String
    languageRnd = Int((3 - 1 + 1) * rnd + 1)
    Select Case languageRnd
        Case 1:
            objOption.language = "DE"
        Case 2:
            objOption.language = "EN"
        Case 3:
            objOption.language = "BG"
    End Select
    Call GetTextComponents
    Call GetAnimalGrammar
    Debug.Print vbTab & vbTab & "language: " & objOption.language

    Dim coloursRnd As String
    coloursRnd = Int((7 - 1 + 1) * rnd + 1)
    Select Case coloursRnd
        Case 1:
            g_strColourMode = "STANDARD"
        Case 2:
            g_strColourMode = "POPART"
        Case 3:
            g_strColourMode = "LSD"
        Case 4:
            g_strColourMode = "SMARTIES"
        Case 5:
            g_strColourMode = "DARKMODE"
        Case 6:
            g_strColourMode = "TV1960"
        Case 7:
            g_strColourMode = "24H"
            frmColourMode.scr24h.Value = Int((4 - (-4) + 1) * rnd + (-4))
            objOption.DAYLIGHT = frmColourMode.scr24h.Value
            objOption.COL_BACK = objOption.DAYLIGHT_COL
            objOption.COL_TEXT = vbBlack
            objOption.COL_RANKINGS = objOption.DAYLIGHT_COL
            Debug.Print vbTab & vbTab & "DAYLIGHT: " & objOption.DAYLIGHT
    End Select
    Debug.Print vbTab & vbTab & "ColourMode: " & g_strColourMode

    objOption.MOMENTUM = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "MOMENTUM: " & objOption.MOMENTUM

    objOption.MOMENTUM_REFRESHRATE = Int((8 - 1 + 1) * rnd + 1)
    Debug.Print vbTab & vbTab & "MOMENTUM_REFRESHRATE: " & objOption.MOMENTUM_REFRESHRATE

    objOption.TACTICS = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "TACTICS: " & objOption.TACTICS
    
    objOption.TACTICS_REVEAL_TAC = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "TACTICS_REVEAL_TAC: " & objOption.TACTICS_REVEAL_TAC
    
    objOption.TACTICS_REVEAL_CURR = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "TACTICS_REVEAL_CURR: " & objOption.TACTICS_REVEAL_CURR
    
    objOption.REFUSE_RUN = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "REFUSE_RUN: " & objOption.REFUSE_RUN
    
    objOption.REFUSAL_RATE = Int((1000 - 10 + 1) * rnd + 10)
    Debug.Print vbTab & vbTab & "REFUSAL_RATE: " & objOption.REFUSAL_RATE
    
    objOption.SLIPSTREAM = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "SLIPSTREAM: " & objOption.SLIPSTREAM
    
    objOption.SLIPSTREAM_DBL = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "SLIPSTREAM_DBL: " & objOption.SLIPSTREAM_DBL
    
    objOption.SLIPSTREAM_SHOW = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "SLIPSTREAM_SHOW: " & objOption.SLIPSTREAM_SHOW

    objOption.FOCUSED_RUN = Int((2 - 0 + 1) * rnd + 0)
    Debug.Print vbTab & vbTab & "FOCUSED_RUN: " & objOption.FOCUSED_RUN
    If objOption.FOCUSED_RUN = 1 Then 'Horse in focus
        objOption.FOCUSED_NR = Int((objRace.NUMBER_STARTING - 1 + 1) * rnd + 1)
        Debug.Print vbTab & vbTab & "FOCUSED_NR: " & objOption.FOCUSED_NR
    End If

    objOption.HIGHLIGHT_FOC = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "HIGHLIGHT_FOC: " & objOption.HIGHLIGHT_FOC

    objOption.METRES_DISPLAY = Int((20 - 1 + 1) * rnd + 1) * 50
    Debug.Print vbTab & vbTab & "METRES_DISPLAY: " & objOption.METRES_DISPLAY

    objOption.TRIBUNES = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "TRIBUNES: " & objOption.TRIBUNES
    
    objOption.SPECTATORS = Int((100 - 0 + 1) * rnd + 0)
    Debug.Print vbTab & vbTab & "SPECTATORS: " & objOption.SPECTATORS
    
    objOption.HOOFPRINTS = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "HOOFPRINTS: " & objOption.HOOFPRINTS
    
    objOption.NAMES_LEFT = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "NAMES_LEFT: " & objOption.NAMES_LEFT
    
    objOption.COLOURS_LEFT = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "COLOURS_LEFT: " & objOption.COLOURS_LEFT
    
    objOption.HIGHLIGHT_FAV = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "HIGHLIGHT_FAV: " & objOption.HIGHLIGHT_FAV
    
    objOption.NAMES_FINISH = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "NAMES_FINISH: " & objOption.NAMES_FINISH

    objOption.NAMES_PHOTO = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "NAMES_PHOTO: " & objOption.NAMES_PHOTO
    
    objOption.PHOTO_BW = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "PHOTO_BW: " & objOption.PHOTO_BW
    
    objOption.RANKING_COL = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "RANKING_COL: " & objOption.RANKING_COL
    
    objOption.RANKING_DELAY = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "RANKING_DELAY: " & objOption.RANKING_DELAY

    objOption.RACE_INFO = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "RACE_INFO: " & objOption.RACE_INFO
    
    If Int((1 - 0 + 1) * rnd + 0) = 1 Then
        objOption.RACE_INFO_POP = False
        objOption.RACE_INFO_WKS = True
        Debug.Print vbTab & vbTab & "RACE_INFO_WKS"
    Else
        objOption.RACE_INFO_POP = True
        objOption.RACE_INFO_WKS = False
        Debug.Print vbTab & vbTab & "RACE_INFO_POP"
    End If
    
    objOption.RACE_INFO_LEADER = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "RACE_INFO_LEADER: " & objOption.RACE_INFO_LEADER
    
    objOption.RACE_INFO_PROGRESS = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "RACE_INFO_PROGRESS: " & objOption.RACE_INFO_PROGRESS
    
    objOption.RACE_INFO_COL_B = Int((16777215 - 0 + 1) * rnd + 0)
    Debug.Print vbTab & vbTab & "RACE_INFO_COL_B: " & objOption.RACE_INFO_COL_B
    
    objOption.RACE_INFO_COL_F = Int((16777215 - 0 + 1) * rnd + 0)
    Debug.Print vbTab & vbTab & "RACE_INFO_COL_F: " & objOption.RACE_INFO_COL_F

    objOption.SPEECH = Int((1 - 0 + 1) * rnd + 0) - 1
    Debug.Print vbTab & vbTab & "SPEECH: " & objOption.SPEECH
    
    objOption.SPEED_FACTOR = Int((5 - 1 + 1) * rnd + 1)
    Debug.Print vbTab & vbTab & "SPEED_FACTOR: " & objOption.SPEED_FACTOR

    'Race specific parameters
    objOption.PARTICULATES_SLIDER = Int((5 - 0 + 1) * rnd + 0)
    Debug.Print vbTab & vbTab & "PARTICULATES_SLIDER: " & objOption.PARTICULATES_SLIDER

    objOption.TIDE = Int((10 - 0 + 1) * rnd + 0)
    Debug.Print vbTab & vbTab & "TIDE: " & objOption.TIDE
    
    objOption.LUGWORMS = Int((100 - 0 + 1) * rnd + 0)
    Debug.Print vbTab & vbTab & "LUGWORMS: " & objOption.LUGWORMS

    objOption.SPACE_PLANET = Int((4 - 0 + 1) * rnd + 0)
    Debug.Print vbTab & vbTab & "SPACE_PLANET: " & objOption.SPACE_PLANET
    
    objOption.SPACE_ALIENS = Int((1 - 0 + 1) * rnd + 0)
    Debug.Print vbTab & vbTab & "SPACE_ALIENS: " & objOption.SPACE_ALIENS
    
    objOption.SPACE_KIDNAPPINGRATE = Int((3 - 1 + 1) * rnd + 1)
    Debug.Print vbTab & vbTab & "SPACE_KIDNAPPINGRATE: " & objOption.SPACE_KIDNAPPINGRATE

End Sub

Private Sub RememberRaceOptions()
    'General parameters
    tempSkipDelay = g_skipDelay
    tempLanguage = objOption.language
    tempColourMode = g_strColourMode
    tempDaylight = frmColourMode.scr24h.Value
    
    'Race parameters
    tempSPEED_FACTOR = objOption.SPEED_FACTOR
    tempSPEECH = objOption.SPEECH
    tempMOMENTUM = objOption.MOMENTUM
    tempMOMENTUM_REFRESHRATE = objOption.MOMENTUM_REFRESHRATE
    tempTACTICS = objOption.TACTICS
    tempTACTICS_REVEAL_TAC = objOption.TACTICS_REVEAL_TAC
    tempTACTICS_REVEAL_CURR = objOption.TACTICS_REVEAL_TAC
    tempREFUSE_RUN = objOption.REFUSE_RUN
    tempREFUSAL_RATE = objOption.REFUSAL_RATE
    tempSLIPSTREAM = objOption.SLIPSTREAM
    tempSLIPSTREAM_DBL = objOption.SLIPSTREAM_DBL
    tempSLIPSTREAM_SHOW = objOption.SLIPSTREAM_SHOW
    tempFOCUSED_RUN = objOption.FOCUSED_RUN
    tempHIGHLIGHT_FOC = objOption.HIGHLIGHT_FOC
    tempMETRES_DISPLAY = objOption.METRES_DISPLAY
    tempTRIBUNES = objOption.TRIBUNES
    tempSPECTATORS = objOption.SPECTATORS
    tempHOOFPRINTS = objOption.HOOFPRINTS
    tempNAMES_LEFT = objOption.NAMES_LEFT
    tempCOLOURS_LEFT = objOption.COLOURS_LEFT
    tempHIGHLIGHT_FAV = objOption.HIGHLIGHT_FAV
    tempNAMES_FINISH = objOption.NAMES_FINISH
    tempNAMES_PHOTO = objOption.NAMES_PHOTO
    tempPHOTO_BW = objOption.PHOTO_BW
    tempRANKING_COL = objOption.RANKING_COL
    tempRANKING_DELAY = objOption.RANKING_DELAY
    tempRACE_INFO = objOption.RACE_INFO
    tempRACE_INFO_POP = objOption.RACE_INFO_POP
    tempRACE_INFO_WKS = objOption.RACE_INFO_WKS
    tempRACE_INFO_LEADER = objOption.RACE_INFO_LEADER
    tempRACE_INFO_PROGRESS = objOption.RACE_INFO_PROGRESS
    tempRACE_INFO_COL_B = objOption.RACE_INFO_COL_B
    tempRACE_INFO_COL_F = objOption.RACE_INFO_COL_F

    'Race specific parameters
    tempPARTICULATES_SLIDER = objOption.PARTICULATES_SLIDER
    tempTIDE = objOption.TIDE
    tempLUGWORMS = objOption.LUGWORMS

End Sub

Private Sub RestoreRaceOptions()

    'General parameters
    g_skipDelay = tempSkipDelay
    g_strColourMode = tempColourMode
    
    If g_strColourMode = "24H" Then
        frmColourMode.scr24h.Value = tempDaylight
        objOption.DAYLIGHT = frmColourMode.scr24h.Value
        objOption.COL_BACK = objOption.DAYLIGHT_COL
    End If
    
    If g_strPlayMode = "RS" Then
        Call RS_StartScreen
        Call RS_InactivateCommandButtons
        objRace.STARTED = False
    End If
    
    objOption.language = tempLanguage
        Call GetTextComponents
        Call GetAnimalGrammar
    
    'Race parameters
    objOption.SPEED_FACTOR = tempSPEED_FACTOR
    objOption.SPEECH = tempSPEECH
    objOption.MOMENTUM = tempMOMENTUM
    objOption.MOMENTUM_REFRESHRATE = tempMOMENTUM_REFRESHRATE
    objOption.TACTICS = tempTACTICS
    objOption.TACTICS_REVEAL_TAC = tempTACTICS_REVEAL_TAC
    objOption.TACTICS_REVEAL_TAC = tempTACTICS_REVEAL_CURR
    objOption.REFUSE_RUN = tempREFUSE_RUN
    objOption.REFUSAL_RATE = tempREFUSAL_RATE
    objOption.SLIPSTREAM = tempSLIPSTREAM
    objOption.SLIPSTREAM_DBL = tempSLIPSTREAM_DBL
    objOption.SLIPSTREAM_SHOW = tempSLIPSTREAM_SHOW
    objOption.FOCUSED_RUN = tempFOCUSED_RUN
    objOption.HIGHLIGHT_FOC = tempHIGHLIGHT_FOC
    objOption.METRES_DISPLAY = tempMETRES_DISPLAY
    objOption.TRIBUNES = tempTRIBUNES
    objOption.SPECTATORS = tempSPECTATORS
    objOption.HOOFPRINTS = tempHOOFPRINTS
    objOption.NAMES_LEFT = tempNAMES_LEFT
    objOption.COLOURS_LEFT = tempCOLOURS_LEFT
    objOption.HIGHLIGHT_FAV = tempHIGHLIGHT_FAV
    objOption.NAMES_FINISH = tempNAMES_FINISH
    objOption.NAMES_PHOTO = tempNAMES_PHOTO
    objOption.PHOTO_BW = tempPHOTO_BW
    objOption.RANKING_COL = tempRANKING_COL
    objOption.RANKING_DELAY = tempRANKING_DELAY
    objOption.RACE_INFO = tempRACE_INFO
    objOption.RACE_INFO_POP = tempRACE_INFO_POP
    objOption.RACE_INFO_WKS = tempRACE_INFO_WKS
    objOption.RACE_INFO_LEADER = tempRACE_INFO_LEADER
    objOption.RACE_INFO_PROGRESS = tempRACE_INFO_PROGRESS
    objOption.RACE_INFO_COL_B = tempRACE_INFO_COL_B
    objOption.RACE_INFO_COL_F = tempRACE_INFO_COL_F
    
    'Race specific parameters
    objOption.PARTICULATES_SLIDER = tempPARTICULATES_SLIDER
    objOption.TIDE = tempTIDE
    objOption.LUGWORMS = tempLUGWORMS

End Sub
