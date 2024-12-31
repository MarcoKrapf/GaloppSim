Attribute VB_Name = "basDeveloperTools"
Option Explicit
Option Private Module

'This module contains non-productive procedures
'which can be used for automatic testing and machine learning
'   Module basDeveloperTools

    'General parameters
    Dim tempSkipDelay As Boolean
    Dim tempLanguage As String
    Dim tempColourMode As String
    Dim tempDaylight As Integer
    
    'Race parameters
    Dim tempSPEED_FACTOR As Integer
    Dim tempSPEECH As Boolean
    Dim tempMOMENTUM_BARS As Boolean
    Dim tempMOMENTUM_ICONS As Boolean
    Dim tempMOMENTUM_REFRESHRATE As Integer
    Dim tempSPEEDMONITOR As Boolean
    Dim tempRSMON_SPEED As Boolean
    Dim tempRSMON_DISTANCE As Boolean
    Dim tempSPEEDMON_REFRESHRATE As Integer
    Dim tempTACTICS As Boolean
    Dim tempTACTICS_REVEAL_TAC As Boolean
    Dim tempTACTICS_REVEAL_CURR As Boolean
    Dim tempREFUSE_RUN As Boolean
    Dim tempREFUSAL_RATE As Integer
    Dim tempSLIPSTREAM_IMPACT As Integer
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
    Dim tempSTARTING_GRID_IN As Boolean
    Dim tempSTARTING_GRID_BEHIND As Boolean
    
    'Race specific parameters
    Dim tempPARTICULATES_SLIDER As Integer
    Dim tempTIDE As Integer
    Dim tempLUGWORMS As Integer
    Dim tempWATER_SPLASHES As Boolean

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
        
        Debug.Print vbTab & "> Test " & TestCase & "/" & colScope.count & " - " & g_wksRaceData.name
        
        Call basAuxiliary.Scroll(1, 1) 'Scroll to the upper left

        If blnRnd Then
            Call TestAutomationRandomSettings
            If blnCurrentSpeed Then objOption.SPEED_FACTOR = tempSPEED_FACTOR
            If blnNoSpeech Then objOption.SPEECH = False
            If blnNoMomentum Then
                objOption.MOMENTUM_BARS = False
                objOption.MOMENTUM_ICONS = False
            End If
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
        Call ShowWinnerPhoto(True, True)
        Application.Wait (Now + TimeValue("0:00:02"))
        Call ShowFinishPhoto(True)
        Application.Wait (Now + TimeValue("0:00:02"))
        Call basAuxiliary.Freeze(0, 0, False) 'Unfreeze
        
        Debug.Print vbTab & "Test " & TestCase & "/" & colScope.count & " finished <" & vbNewLine

    Next
    
    Call basAuxiliary.Scroll(1, 1) 'Scroll to the upper left
    Call RS_MenuAreaShow
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
    languageRnd = Int((3 - 1 + 1) * Rnd + 1)
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
    coloursRnd = Int((7 - 1 + 1) * Rnd + 1)
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
            frmColourMode.scr24h.Value = Int((4 - (-4) + 1) * Rnd + (-4))
            objOption.DAYLIGHT = frmColourMode.scr24h.Value
            objOption.COL_BACK = objOption.DAYLIGHT_COL
            objOption.COL_TEXT = vbBlack
            objOption.COL_RANKINGS = objOption.DAYLIGHT_COL
            Debug.Print vbTab & vbTab & "DAYLIGHT: " & objOption.DAYLIGHT
    End Select
    Debug.Print vbTab & vbTab & "ColourMode: " & g_strColourMode

    objOption.MOMENTUM_BARS = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "MOMENTUM_BARS: " & objOption.MOMENTUM_BARS
    
    objOption.MOMENTUM_ICONS = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "MOMENTUM_ICONS: " & objOption.MOMENTUM_ICONS

    objOption.MOMENTUM_REFRESHRATE = Int((60 - 1 + 1) * Rnd + 1)
    Debug.Print vbTab & vbTab & "MOMENTUM_REFRESHRATE: " & objOption.MOMENTUM_REFRESHRATE
    
    objOption.SPEEDMONITOR = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "SPEEDMONITOR: " & objOption.SPEEDMONITOR
    
    If Int((1 - 0 + 1) * Rnd + 0) = 1 Then
        objOption.RSMON_SPEED = False
        objOption.RSMON_DISTANCE = True
        Debug.Print vbTab & vbTab & "RSMON_DISTANCE"
    Else
        objOption.RSMON_SPEED = True
        objOption.RSMON_DISTANCE = False
        Debug.Print vbTab & vbTab & "RSMON_SPEED"
    End If

    objOption.SPEEDMON_REFRESHRATE = Int((50 - 10 + 1) * Rnd + 10)
    Debug.Print vbTab & vbTab & "SPEEDMON_REFRESHRATE: " & objOption.SPEEDMON_REFRESHRATE
    
    If objOption.SPEEDMONITOR = 0 Then
        'Reset the RSMon collection
        Set g_colRSMon = Nothing
        Set g_colRSMon = New Collection
        'Assign between 1 and 3 horses
        Dim nr As Integer
        For nr = 1 To Int((3 - 1 + 1) * Rnd + 1)
            g_colRSMon.Add nr
        Next
    End If

    objOption.TACTICS = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "TACTICS: " & objOption.TACTICS
    
    objOption.TACTICS_REVEAL_TAC = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "TACTICS_REVEAL_TAC: " & objOption.TACTICS_REVEAL_TAC
    
    objOption.TACTICS_REVEAL_CURR = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "TACTICS_REVEAL_CURR: " & objOption.TACTICS_REVEAL_CURR
    
    objOption.REFUSE_RUN = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "REFUSE_RUN: " & objOption.REFUSE_RUN
    
    objOption.REFUSAL_RATE = Int((1000 - 10 + 1) * Rnd + 10)
    Debug.Print vbTab & vbTab & "REFUSAL_RATE: " & objOption.REFUSAL_RATE
    
    objOption.SLIPSTREAM_IMPACT = Int((2 - 0 + 1) * Rnd + 0)
    Debug.Print vbTab & vbTab & "SLIPSTREAM_IMPACT: " & objOption.SLIPSTREAM_IMPACT
    
    objOption.SLIPSTREAM_SHOW = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "SLIPSTREAM_SHOW: " & objOption.SLIPSTREAM_SHOW

    objOption.FOCUSED_RUN = Int((2 - 0 + 1) * Rnd + 0)
    Debug.Print vbTab & vbTab & "FOCUSED_RUN: " & objOption.FOCUSED_RUN
    If objOption.FOCUSED_RUN = 1 Then 'Horse in focus
        objOption.FOCUSED_NR = Int((objRace.NUMBER_STARTING - 1 + 1) * Rnd + 1)
        Debug.Print vbTab & vbTab & "FOCUSED_NR: " & objOption.FOCUSED_NR
    End If

    objOption.HIGHLIGHT_FOC = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "HIGHLIGHT_FOC: " & objOption.HIGHLIGHT_FOC

    objOption.METRES_DISPLAY = Int((20 - 1 + 1) * Rnd + 1) * 50
    Debug.Print vbTab & vbTab & "METRES_DISPLAY: " & objOption.METRES_DISPLAY

    objOption.TRIBUNES = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "TRIBUNES: " & objOption.TRIBUNES
    
    objOption.SPECTATORS = Int((100 - 0 + 1) * Rnd + 0)
    Debug.Print vbTab & vbTab & "SPECTATORS: " & objOption.SPECTATORS
    
    objOption.HOOFPRINTS = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "HOOFPRINTS: " & objOption.HOOFPRINTS
    
    objOption.NAMES_LEFT = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "NAMES_LEFT: " & objOption.NAMES_LEFT
    
    objOption.COLOURS_LEFT = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "COLOURS_LEFT: " & objOption.COLOURS_LEFT
    
    objOption.HIGHLIGHT_FAV = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "HIGHLIGHT_FAV: " & objOption.HIGHLIGHT_FAV
    
    objOption.NAMES_FINISH = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "NAMES_FINISH: " & objOption.NAMES_FINISH

    objOption.NAMES_PHOTO = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "NAMES_PHOTO: " & objOption.NAMES_PHOTO
    
    objOption.PHOTO_BW = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "PHOTO_BW: " & objOption.PHOTO_BW
    
    objOption.RANKING_COL = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "RANKING_COL: " & objOption.RANKING_COL
    
    objOption.RANKING_DELAY = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "RANKING_DELAY: " & objOption.RANKING_DELAY

    objOption.RACE_INFO = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "RACE_INFO: " & objOption.RACE_INFO
    
    If Int((1 - 0 + 1) * Rnd + 0) = 1 Then
        objOption.RACE_INFO_POP = False
        objOption.RACE_INFO_WKS = True
        Debug.Print vbTab & vbTab & "RACE_INFO_WKS"
    Else
        objOption.RACE_INFO_POP = True
        objOption.RACE_INFO_WKS = False
        Debug.Print vbTab & vbTab & "RACE_INFO_POP"
    End If
    
    objOption.RACE_INFO_LEADER = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "RACE_INFO_LEADER: " & objOption.RACE_INFO_LEADER
    
    objOption.RACE_INFO_PROGRESS = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "RACE_INFO_PROGRESS: " & objOption.RACE_INFO_PROGRESS
    
    objOption.RACE_INFO_COL_B = Int((16777215 - 0 + 1) * Rnd + 0)
    Debug.Print vbTab & vbTab & "RACE_INFO_COL_B: " & objOption.RACE_INFO_COL_B
    
    objOption.RACE_INFO_COL_F = Int((16777215 - 0 + 1) * Rnd + 0)
    Debug.Print vbTab & vbTab & "RACE_INFO_COL_F: " & objOption.RACE_INFO_COL_F

    objOption.SPEECH = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "SPEECH: " & objOption.SPEECH
    
    objOption.SPEED_FACTOR = Int((5 - 1 + 1) * Rnd + 1)
    Debug.Print vbTab & vbTab & "SPEED_FACTOR: " & objOption.SPEED_FACTOR

    Debug.Print vbTab & vbTab & "STARTING PROCEDURE: "
    If Int((1 - 0 + 1) * Rnd + 0) = 1 Then
        objOption.STARTING_GRID_BEHIND = False
        objOption.STARTING_GRID_IN = True
        Debug.Print vbTab & vbTab & "STARTING_GRID_IN"
    Else
        objOption.STARTING_GRID_BEHIND = True
        objOption.STARTING_GRID_IN = False
        Debug.Print vbTab & vbTab & "STARTING_GRID_BEHIND"
    End If

    'Race specific parameters
    objOption.PARTICULATES_SLIDER = Int((5 - 0 + 1) * Rnd + 0)
    Debug.Print vbTab & vbTab & "PARTICULATES_SLIDER: " & objOption.PARTICULATES_SLIDER

    objOption.TIDE = Int((10 - 0 + 1) * Rnd + 0)
    Debug.Print vbTab & vbTab & "TIDE: " & objOption.TIDE
    
    objOption.LUGWORMS = Int((100 - 0 + 1) * Rnd + 0)
    Debug.Print vbTab & vbTab & "LUGWORMS: " & objOption.LUGWORMS
    
    objOption.WATER_SPLASHES = Int((1 - 0 + 1) * Rnd + 0) - 1
    Debug.Print vbTab & vbTab & "WATER_SPLASHES: " & objOption.WATER_SPLASHES

    objOption.SPACE_PLANET = Int((4 - 0 + 1) * Rnd + 0)
    Debug.Print vbTab & vbTab & "SPACE_PLANET: " & objOption.SPACE_PLANET
    
    objOption.SPACE_ALIENS = Int((1 - 0 + 1) * Rnd + 0)
    Debug.Print vbTab & vbTab & "SPACE_ALIENS: " & objOption.SPACE_ALIENS
    
    objOption.SPACE_KIDNAPPINGRATE = Int((3 - 1 + 1) * Rnd + 1)
    Debug.Print vbTab & vbTab & "SPACE_KIDNAPPINGRATE: " & objOption.SPACE_KIDNAPPINGRATE

End Sub

Public Sub MachineLearningSimulation(colScope As Collection, lngRepetitions As Long, varSave As Variant)

    Dim intRace As Integer 'Current race
    Dim lngRep As Long 'Current repetition
    Dim lngSimCase As Long 'Current simulation case
    Dim timeStart As Date 'For calculating the execution time

    Call RememberRaceOptions
    g_skipDelay = True
    timeStart = Now
    
    Debug.Print vbNewLine & vbNewLine & ">>> START MACHINE LEARNING SIMULATION" & vbNewLine
    frmMachineLearning.lblMLProgress.caption = ""
    
    If frmMachineLearning.chkExport Then
    
        Dim intFileNr As Integer 'Output channel
        intFileNr = FreeFile 'Assign the next free number
        
        Debug.Print "  Save to " & varSave
        
        'Check whether the file exists
        If Dir(varSave) <> "" Then 'Yes
            Debug.Print "  --> add to existing file" & vbNewLine
        Else 'No
            Debug.Print "  --> create new file" & vbNewLine
            Open varSave For Output As #intFileNr
            Print #intFileNr, "ID" & ";" _
                    & "RACE_ID" & ";" _
                    & "PARTICIPANTS" & ";" _
                    & "METRES" & ";" _
                    & "ENROLLED" & ";" _
                    & "STARTERS" & ";" _
                    & "STARTING_ORDER" & ";" _
                    & "RACETYPE" & ";" _
                    & "REFUSE_TO_RUN" & ";" _
                    & "REFUSAL_RATE" & ";" _
                    & "SLIPSTREAM_IMPACT" & ";" _
                    & "NAME_AND_STARTING_NR" & ";" _
                    & "STATUS" & ";" _
                    & "PLACEMENT" & ";" _
                    & "BOX_NR" & ";" _
                    & "BASIC_SPEED" & ";" _
                    & "FORM_FACTOR" & ";" _
                    & "TACTICS"
            Close #intFileNr
        End If
        
    End If
    
    'Simulate races without graphical representation
    For intRace = 1 To colScope.count 'Loop through the simulation scope
        Set g_wksRaceData = ThisWorkbook.Worksheets(colScope(intRace)) 'Get the next race

        For lngRep = 1 To lngRepetitions
            lngSimCase = lngSimCase + 1
            
            If frmMachineLearning.chkDebug Then
                    Debug.Print "  > Simulation case " & lngSimCase & "/" & (colScope.count * lngRepetitions) & " starting - " & g_wksRaceData.name & vbNewLine
            End If
            
            frmMachineLearning.frameMLprogress.caption = "SIMULATION RACE PROGRESS: " & lngSimCase & "/" & colScope.count * lngRepetitions
            With frmMachineLearning.lblMLProgress
                .width = 348 / (colScope.count * lngRepetitions) * lngSimCase
                .BackColor = vbBlue
            End With
    
            Call basMainCode.ML_NewRace(lngSimCase, varSave) 'Start a Machine Learning simulation race
    
            If frmMachineLearning.chkDebug Then
                    Debug.Print vbNewLine & "  Simulation case " & lngSimCase & "/" & (colScope.count * lngRepetitions) & " finished <" & vbNewLine
            End If
    
        Next lngRep
    Next

    Call RestoreRaceOptions
    
    With frmMachineLearning.lblMLProgress
        .caption = "FINISHED"
        .BackColor = vbGreen
    End With
    
    Debug.Print "EXECUTION TIME: " & Format(Now - timeStart, "HH:MM:SS")
    Debug.Print "MACHINE LEARNING SIMULATION FINISHED <<<" & vbNewLine & vbNewLine
    
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
    tempMOMENTUM_BARS = objOption.MOMENTUM_BARS
    tempMOMENTUM_ICONS = objOption.MOMENTUM_ICONS
    tempMOMENTUM_REFRESHRATE = objOption.MOMENTUM_REFRESHRATE
    tempSPEEDMONITOR = objOption.SPEEDMONITOR
    tempRSMON_SPEED = objOption.RSMON_SPEED
    tempRSMON_DISTANCE = objOption.RSMON_DISTANCE
    tempSPEEDMON_REFRESHRATE = objOption.SPEEDMON_REFRESHRATE
    tempTACTICS = objOption.TACTICS
    tempTACTICS_REVEAL_TAC = objOption.TACTICS_REVEAL_TAC
    tempTACTICS_REVEAL_CURR = objOption.TACTICS_REVEAL_TAC
    tempREFUSE_RUN = objOption.REFUSE_RUN
    tempREFUSAL_RATE = objOption.REFUSAL_RATE
    tempSLIPSTREAM_IMPACT = objOption.SLIPSTREAM_IMPACT
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
    tempSTARTING_GRID_IN = objOption.STARTING_GRID_IN
    tempSTARTING_GRID_BEHIND = objOption.STARTING_GRID_BEHIND

    'Race specific parameters
    tempPARTICULATES_SLIDER = objOption.PARTICULATES_SLIDER
    tempTIDE = objOption.TIDE
    tempLUGWORMS = objOption.LUGWORMS
    tempWATER_SPLASHES = objOption.WATER_SPLASHES

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
    objOption.MOMENTUM_BARS = tempMOMENTUM_BARS
    objOption.MOMENTUM_ICONS = tempMOMENTUM_ICONS
    objOption.MOMENTUM_REFRESHRATE = tempMOMENTUM_REFRESHRATE
    objOption.SPEEDMONITOR = tempSPEEDMONITOR
    objOption.RSMON_SPEED = tempRSMON_SPEED
    objOption.RSMON_DISTANCE = tempRSMON_DISTANCE
    objOption.SPEEDMON_REFRESHRATE = tempSPEEDMON_REFRESHRATE
    objOption.TACTICS = tempTACTICS
    objOption.TACTICS_REVEAL_TAC = tempTACTICS_REVEAL_TAC
    objOption.TACTICS_REVEAL_TAC = tempTACTICS_REVEAL_CURR
    objOption.REFUSE_RUN = tempREFUSE_RUN
    objOption.REFUSAL_RATE = tempREFUSAL_RATE
    objOption.SLIPSTREAM_IMPACT = tempSLIPSTREAM_IMPACT
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
    objOption.STARTING_GRID_IN = tempSTARTING_GRID_IN
    objOption.STARTING_GRID_BEHIND = tempSTARTING_GRID_BEHIND
    
    'Race specific parameters
    objOption.PARTICULATES_SLIDER = tempPARTICULATES_SLIDER
    objOption.TIDE = tempTIDE
    objOption.LUGWORMS = tempLUGWORMS
    objOption.WATER_SPLASHES = tempWATER_SPLASHES

End Sub
