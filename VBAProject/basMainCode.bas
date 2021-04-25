Attribute VB_Name = "basMainCode"
Option Explicit 'Force variable declaration
Option Private Module 'Prevent the public procedures in this module from being accessed from outside this module

'This module contains the main code of the horse racing simulator with most of the logic
'   Module basMainCode


'GALOPPSIM - Version 151.50 (25 April 2021)
'Horse Racing Simulator for Microsoft Excel
'Author: Marco Matjes
'info@galoppsim.racing - https://galoppsim.racing/
'License: GNU General Public License v3.0
            

'NAMING CONVENTIONS
'------------------
    'Scope prefixes
        'g_ = global (project-wide)
        'm_ = module-level
    'Excel elements
        'wks = WorkSheet
        'bas = Standard module
        'cls = Class module
        'frm = UserForm
    'UserForm elements
        'lbl = Label
        'txt = TextBox
        'img = Image
        'chk = CheckBox
        'opt = OptionButton
        'cmd = CommandButton
        'cbo = ComboBox
        'lst = ListBox
    'Constants
        'c_
    'Variable data types
        'str = String
        'int = Integer
        'lng = Long
        'dbl = Double
        'bln = Boolean
        'var = Variant
        'col = Collection
        'obj = Object
        'ole = OLEobject
        'arr_ = Array
        'enum = Enumeration


'GLOBAL CONSTANTS AND VARIABLES
'------------------------------
Public Const g_c_tool As String = "GaloppSim" 'Name of the tool
Public Const g_c_version As String = "151.50" 'Version of the tool
Public Const g_c_email As String = "info@galoppsim.racing" 'Contact e-mail address
Public Const g_c_defaultRaceOptionsFile As String = "RACEOPTIONS" 'File name for race options
Public Const g_c_defaultFileType As String = ".GALOPPSIM" 'File type for GaloppSim files
Public Const g_c_errorLogFileName As String = "GALOPPSIM_ERRORLOG" 'File name for error logging
Public g_MLdataFileName As String 'File name for Machine Learning export
Public g_defaultPath As String 'Path for GaloppSim files (Add-in standard folder)
Public g_defaultMLpath As String 'Path for the Machine Learning files
Public g_defaultAutoSavePath As String 'Path for the auto-save function after a race
Public g_strPlayMode As String ' "AI" = AddIn (.xlam) / "RS" = Run Simple (.xlsm)
Public g_RibbonGaloppSim As IRibbonUI 'Custom ribbon (only used for the AI edition)
Public AI_started As Boolean 'Will be set true in AI edition when activating the GALOPPSIM menu tab the first time
Public g_skipDelay As Boolean 'For skipping delay commands (Application.Wait)
Public g_errorLogPath As String 'Path for error logging
Public g_errorLogging As Boolean 'Error logging on/off
Public g_payoutLogging As Boolean 'Pay-out logging for bets on/off
Public objBasicData As clsBasicData 'Object for all kind of basic data
Public objRace As clsRace 'Object for all the data of a race
Public objOption As clsOptions 'Object for the race and excel options
Public objSpeed As clsSpeed 'Object for speed data
Public objStat As clsStatisctics 'Object for statistical data
Public g_colRSbuttons As Collection 'Menu buttons in the RS edition
Public g_arr_varHorses() As Variant 'All information about the horses
Public g_arr_varReplay_RaceData() As Variant 'Race data for replaying a race
Public g_arr_varReplay_HorseData() As Variant 'Horse data for replaying a race
Public g_colRacesInstalled As Collection 'List of installed races
Public g_oleComboRaces As OLEObject 'ComboBox with installed races in the RS edition
Public g_colBetSlips As Collection 'List with all betting slips
Public g_arr_Text() As Variant 'All text components
Public g_arr_Grammar(1 To 8) As String 'All animal grammar components
Public g_objLabel As MSForms.Label 'For creating Labels on UserForms at runtime
Public g_objDropdown As MSForms.ComboBox 'For creating ComboBoxes on UserForms at runtime
Public g_objListbox As MSForms.ListBox 'For creating ListBoxes on UserForms at runtime
Public g_strInpBoxReturnValue As String 'Return value of an input box
Public g_enumButton As String 'Return value of the pressed button
Public g_shpBar As Shape 'Race progress bar on the worksheet
Public g_shpFrame As Shape 'Frame around a race progress bar on the worksheet
Public g_strColourMode As String 'Colour mode (Standard, PopArt etc.)
Public g_colRSMon As Collection 'Selected horses for Race Speed Monitor
Public g_arr_Developer(1 To 3) As Byte 'Super Developer Tools

'Existing Worksheets with data
Public g_wksTEXT As Worksheet 'Worksheet with text components
Public g_wksPIC As Worksheet 'Worksheet with picture data
Public g_wksTEC As Worksheet 'Worksheet with technical data (speed, tactics...)
Public g_wksTCASE As Worksheet 'Worksheet for editing the test case repository

'Worksheets created at runtime
Public g_wksRace As Worksheet 'Worksheet for the race
Public g_wksRaceData As Worksheet 'Worksheet with the race data
Public g_wksMovie As Worksheet 'Worksheet for the movie


'GLOBAL ENUMERATIONS
'-------------------

Public Enum enumButton 'Buttons in pop-ups
    OK
    CancelOK
    YesNo
    Cancel
    yes
    no
End Enum

Public Enum enumCamera 'Camera mode
    standard
    focus_horse
    focus_leader
End Enum

Public Enum enumPlanets 'Planets in a space race
    moon
    mars
    jupiter
    pluto
    saturn
End Enum

Public Enum enumAliens 'Alien behaviour in a space race
    friendly
    unfriendly
End Enum

Public Enum enumBetType 'Bet type
    win
    show
    exacta
    x2sur4
    trifecta
    superfecta
End Enum


'VARIABLES AND CONSTANTS ON MODULE-LEVEL
'---------------------------------------
Dim m_wksCheck As Worksheet 'Variable to check whether the worksheet GALOPPSIM already exists
Dim m_arr_varPhotofinish() As Variant 'Position of each horse on the photo of the finish
Dim m_arr_varResultsCalc() As Variant 'Calculation of the position at the finish line
Dim m_arr_varResults() As Variant 'Race results list
Dim m_arr_varTrackGraphics() As Variant 'Array for track graphics and advertising data
Dim m_intAdvertisingHeight As Integer 'Row height of the advertising area
Dim m_intTrackCellHeight As Integer 'Cell height (race track)
Dim m_dblTrackCellWidth As Double 'Cell width (race track)
Dim m_intFontSize As Integer 'Font size of the horse names and the hoof prints
Dim m_intHorsesRunning As Integer 'Number of horses currently running
Dim m_byteFavourite(1 To 3) As Byte 'Array for three predicted favourites of the race
Dim m_dblFavCalc(1 To 3) As Double 'Array for calculating the favourites
Dim m_intHorsesFinishing As Integer 'Variable that counts how many horses arrive at the finish line in one loop
Dim m_intFinishLoop As Integer 'Variable that counts in which computation period a placement was calculated
Dim m_blnWin As Boolean 'Flag indicating that a horse has won the race
Dim m_blnPhotofinish As Boolean 'Flag indicating whether there is a photo finish
Dim m_blnDeadHeat As Boolean 'Dead heat (more than one horse are at 1st place)
Dim m_intPlace As Integer 'Placement in the finish
Dim m_arr_lngLugworms() As Long 'Array for the lugworm characters in a mudflats race
Dim m_shSpeedChart As Shape 'Chart for the speed lines
Dim m_arr_strSpectators(1) As String
Dim i As Integer, j As Integer, k As Integer, m As Integer 'Counting variables for loops
Dim z As Long 'Auxiliary variable for loops

'VARIABLES USED FOR MATJES GRAND PRIX RACES (GROSSER MATJES-PREIS)
'-----------------------------------------------------------------
Dim GMP_mode As String
Dim GMP_qualified As Integer
Dim CC_qualifed As Integer

'PROCEDURES
'----------

'Main procedure for starting a new race
Public Sub NewRace(Optional test As Boolean)

    On Error GoTo ERRORHANDLING 'In case an error occurs

    'Override settings dependent on the selected race and colour mode
    Call GetColours_colMode
    Call GetColours_specRace
    
    'Close pop-ups if visible
    If frmBettingAnalysis.Visible Then Unload frmBettingAnalysis 'Analysis of the bets
    If frmRS_navigation.Visible Then Unload frmRS_navigation 'Navigation panel (RS edition only)
    
    'Reset the betting slip collection
    Set g_colBetSlips = Nothing
    Set g_colBetSlips = New Collection
    
    If Not test Then Call AssignRaceSheet(False)
    Call GetRaceData(False) 'Get the race data from the worksheet with the selected race
    Call GetAnimalGrammar 'Get grammar components according to the selected language
    Call AssignBasicValues(test) 'Get basic values
    Call GetHorseData 'Get data about the horses according to the selected race
    If Not test Then Call CheckBettingAllowed 'Check whether betting is allowed for this race
    If Not test Then Call CoronavirusTest 'Check whether Coronavirus testing is mandatory
    If Not test Then Call CalculateFavourites 'Calculate the race favourites
    If Not test Then Call ShowStartPopup 'Show the pop-up with the race to run
    If test Then 'Parameters for automatic testing
'        Application.VBE.MainWindow.WindowState = 1 'Minimize the VBA editor
        objRace.STARTED = True
    End If
    
    If objRace.STARTED Then
        If g_strPlayMode = "AI" Then
            Call CreateRaceSheet 'Create the worksheet "GALOPPSIM"
            Call AI_ExcelModeStart
            Call CursorAway 'Place the cursor far away (in the upper right corner of the screen)
        End If
        
        If g_strPlayMode = "RS" Then
            Call basAuxiliary.ActivateRaceSheet
                Cells.Clear 'Clear the whole worksheet
                With Cells(2, 2) 'GaloppSim title and version
                    .Font.name = "Arial Black"
                    .Value = g_c_tool & " (v" & g_c_version & ")"
                End With
            Call RS_MenuAreaHide 'Hide the controls on the worksheet
        End If

        Call DrawRaceTrack(False) 'Race track generation
        Call DrawHorseNames 'Write the horse names on the race track if selected
        If objOption.MOMENTUM_BARS Then Call MomentumFormattings_Bars 'Prepare the race sheet for momentum speed bars
        If objOption.MOMENTUM_ICONS Then Call MomentumFormattings_Icons 'Prepare the race sheet for momentum icons
        If objOption.RACE_INFO And objOption.RACE_INFO_POP Then Call RaceInfoPopup 'Pop-up with race the info if selected
        Call CheckSpaceRace(test) 'Check whether a space race is chosen
        If Not test Then Call RaceWelcome 'Pop-up with a warm welcome to the race
        Call StartingGrid 'Put the horses in the gates
        Call RacePresentation(test, False) 'Presentation of the horses
        Call RunRace(False, test, False) 'Race start (ml=False test=test replay=False)
        Call RankNotFinished 'Find the horses that did not finish
        Call TransmitPlacements 'Write the placements to the main array
        If Not test Then Call RaceFinished 'Info pop-up when the race is over
        Call CheckDeadHeat 'Check whether more than one horse has won
        Call DrawRankingList(True, test, False) 'Race results on the ranking list
        Call DrawWinnerPhoto 'Show a photo of the winner
        If objOption.BET_PLACED And objOption.BET_ANALYSIS Then Call AnalyseBettings 'Pop-up with the bet slips analysis
        
        If g_strPlayMode = "RS" And Not test Then 'Show the navigation panel (RS edition only)
            Call basAuxiliary.ActivateRaceSheet
            'Activate buttons
            With g_wksRace
                .OLEObjects("finishphoto").Object.Enabled = True
                .OLEObjects("results").Object.Enabled = True
                .OLEObjects("winner").Object.Enabled = True
                If objOption.BET_PLACED Then
                    .OLEObjects("bets").Object.Enabled = True
                End If
                .OLEObjects("saverace").Object.Enabled = True
                .OLEObjects("replay").Object.Enabled = True
                .rows(1).RowHeight = 8
                .rows(2).RowHeight = 18
            End With
            Call RS_MenuAreaShow
            frmRS_navigation.show (vbModeless) 'Show a pop-up with the navigation panel
        End If
        
        If g_strPlayMode = "RS" And test Then Call RS_MenuAreaShow

        If g_strPlayMode = "AI" Then
            g_RibbonGaloppSim.Invalidate 'Refresh the status buttons
            Call AI_ExcelModeEnd
        End If
        
    End If

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "NewRace()")
    Call basAuxiliary.CodeCrash
End Sub

'Replay a race
Public Sub RaceReplay()

    On Error GoTo ERRORHANDLING 'In case an error occurs

    objRace.STARTED = True

    'Inactivate some buttons (RS edition only)
    If g_strPlayMode = "RS" Then Call RS_InactivateCommandButtons

    'Close pop-ups if visible
    If frmBettingAnalysis.Visible Then Unload frmBettingAnalysis 'Analysis of the bets
    If frmRS_navigation.Visible Then Unload frmRS_navigation 'Navigation panel (RS edition only)

    'Delete speed chart
    Dim sh As Shape
    For Each sh In g_wksRace.Shapes
       If sh.name = "SpeedChart" Then sh.Delete
    Next
    
    Call AssignRaceSheet(True)
    
    'Get the race data
    objRace.RACE_ID = g_arr_varReplay_RaceData(3, 1)
    objRace.PARTICIPANTS = g_arr_varReplay_RaceData(4, 1)
    objRace.RACE_NAME = g_arr_varReplay_RaceData(5, 1)
    objRace.RACE_YEAR = g_arr_varReplay_RaceData(6, 1)
    objRace.TRACK_LOCATION = g_arr_varReplay_RaceData(7, 1)
    objRace.COUNTRY_CODE = g_arr_varReplay_RaceData(8, 1)
    If g_arr_varReplay_RaceData(8, 1) = "MOON" Then
        Select Case g_arr_varReplay_RaceData(24, 1)
            Case enumPlanets.moon
                objRace.COUNTRY_NAME = GetText(g_arr_Text, "RACESPEC007")
            Case enumPlanets.mars
                objRace.COUNTRY_NAME = GetText(g_arr_Text, "RACESPEC008")
            Case enumPlanets.jupiter
                objRace.COUNTRY_NAME = GetText(g_arr_Text, "RACESPEC009")
            Case enumPlanets.pluto
                objRace.COUNTRY_NAME = GetText(g_arr_Text, "RACESPEC010")
            Case enumPlanets.saturn
                objRace.COUNTRY_NAME = GetText(g_arr_Text, "RACESPEC011")
        End Select
    Else
        objRace.COUNTRY_NAME = GetCountryName(objRace.COUNTRY_CODE, objOption.language)
    End If
    objRace.TRACK_NAME = g_arr_varReplay_RaceData(9, 1)
    objRace.TRACK_COLOUR = g_arr_varReplay_RaceData(10, 1)
    objRace.TRACK_SURFACE = g_arr_varReplay_RaceData(11, 1)
    objRace.RACE_TYPE = g_arr_varReplay_RaceData(12, 1)
    objRace.METRES = g_arr_varReplay_RaceData(13, 1)
    objRace.STARTING_GATE = g_arr_varReplay_RaceData(14, 1)
    objRace.NUMBER_ENROLLED = g_arr_varReplay_RaceData(15, 1)
    objRace.NUMBER_STARTING = g_arr_varReplay_RaceData(16, 1)
    objRace.LANES_FIX_OR_RANDOM = g_arr_varReplay_RaceData(17, 1)
    objRace.ADVERTISING = g_arr_varReplay_RaceData(18, 1)
    objRace.SPECIAL = g_arr_varReplay_RaceData(19, 1)
    objOption.SPECTATORS = g_arr_varReplay_RaceData(23, 1)
    
    Call GetAnimalGrammar 'Get grammar components according to the selected language
    Call AssignBasicValues 'Get basic values

    'Override settings dependent on the selected race and colour mode
    Call GetColours_colMode
    Call GetColours_specRace

    'Get the horse data
    ReDim g_arr_varHorses(1 To UBound(g_arr_varReplay_HorseData), 0 To 31)
    For i = 1 To g_arr_varReplay_RaceData(15, 1) 'NUMBER_ENROLLED
        g_arr_varHorses(i, 0) = g_arr_varReplay_HorseData(i, 0) 'Status before the race
        g_arr_varHorses(i, 1) = g_arr_varReplay_HorseData(i, 1) 'Name of the horse
        g_arr_varHorses(i, 2) = g_arr_varReplay_HorseData(i, 2) 'Horse colour
        g_arr_varHorses(i, 3) = g_arr_varReplay_HorseData(i, 3) 'Row number on which the horse is running
        g_arr_varHorses(i, 4) = objBasicData.LEFT_COLS + 12 'Starting position (column number)
        g_arr_varHorses(i, 9) = 0 'Reset the exact (internal) horse position
        g_arr_varHorses(i, 11) = g_arr_varReplay_HorseData(i, 11) 'Starting number
        g_arr_varHorses(i, 14) = g_arr_varReplay_HorseData(i, 14) 'Race speed log (array with each step)
        g_arr_varHorses(i, 15) = g_arr_varReplay_HorseData(i, 15) 'Starting gate
        Select Case g_arr_varReplay_HorseData(i, 20) 'Convert the current status
            Case "FINISHED", "KIDNAPPED"
                g_arr_varHorses(i, 20) = "RUNNING"
            Case "REFUSED"
                g_arr_varHorses(i, 20) = "REFUSED"
            Case Else
            
        End Select
        g_arr_varHorses(i, 23) = g_arr_varReplay_HorseData(i, 23) 'Picture of the winner
        g_arr_varHorses(i, 24) = g_arr_varReplay_HorseData(i, 24) 'For SPECIAL purposes
        g_arr_varHorses(i, 25) = g_arr_varReplay_HorseData(i, 25) 'Tactics
    Next i
    
    If g_strPlayMode = "AI" Then
        Call CreateRaceSheet 'Create the worksheet "GALOPPSIM"
        Call AI_ExcelModeStart
        Call CursorAway 'Place the cursor far away (in the upper right corner of the screen)
    End If
    
    If g_strPlayMode = "RS" Then
        Call basAuxiliary.ActivateRaceSheet
            Cells.Clear 'Clear the whole worksheet
            With Cells(2, 2) 'GaloppSim title and version
                .Font.name = "Arial Black"
                .Value = g_c_tool & " (v" & g_c_version & ")"
            End With
        Call RS_MenuAreaHide 'Hide the controls on the worksheet
    End If

    Call DrawRaceTrack(True) 'Race track generation
    Call DrawHorseNames 'Write the horse names on the race track if selected
    If objOption.MOMENTUM_BARS Then Call MomentumFormattings_Bars 'Prepare the race sheet for momentum speed bars
    If objOption.MOMENTUM_ICONS Then Call MomentumFormattings_Icons 'Prepare the race sheet for momentum icons
    If objOption.RACE_INFO And objOption.RACE_INFO_POP Then Call RaceInfoPopup 'Pop-up with race the info if selected
    Call CheckSpaceRace 'Check whether a space race is chosen
    If objOption.FOCUSED_RUN = enumCamera.focus_horse Or objOption.SPEEDMONITOR Then frmReplayOptions.show
    
    'Replay information
    Dim messagetext As String
    messagetext = GetText(g_arr_Text, "RACE036") & " " _
        & DateSerial(g_arr_varReplay_RaceData(2, 1)(1), g_arr_varReplay_RaceData(2, 1)(2), g_arr_varReplay_RaceData(2, 1)(3))
    If objOption.SPEECH Then Call SpeechOut(messagetext) 'Voice output if selected
    Call ShowMessagePopup(objRace.RACE_NAME & " " & objRace.RACE_YEAR, messagetext, enumButton.OK, vbModal) 'Show a pop-up

    Call StartingGrid 'Put the horses in the gates
    Call RacePresentation(False, True) 'Presentation of the horses
    Call RunRace(False, False, True) 'Race start (ml=False test=test replay=False)
    Call RankNotFinished 'Find the horses that did not finish
    Call TransmitPlacements 'Write the placements to the main array
    Call CheckDeadHeat 'Check whether more than one horse has won
    Call DrawRankingList(False, False, True) 'Race results on the ranking list
    Call DrawWinnerPhoto 'Show a photo of the winner
    If g_strPlayMode = "RS" Then  'Show the navigation panel (RS edition only)
        Call basAuxiliary.ActivateRaceSheet
        'Activate buttons
        With g_wksRace
            .OLEObjects("finishphoto").Object.Enabled = True
            .OLEObjects("results").Object.Enabled = True
            .OLEObjects("winner").Object.Enabled = True
            .OLEObjects("bets").Object.Enabled = False
            .OLEObjects("replay").Object.Enabled = True
            .rows(1).RowHeight = 8
            .rows(2).RowHeight = 18
        End With
        Call RS_MenuAreaShow
        frmRS_navigation.show (vbModeless) 'Show a pop-up with the navigation panel
    End If

    If g_strPlayMode = "AI" Then
        g_RibbonGaloppSim.Invalidate 'Refresh the status buttons
        Call AI_ExcelModeEnd
    End If

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "Replay()")
    Call basAuxiliary.CodeCrash
End Sub

'Starting procedure (called by the Workbook_Open event in the RS edition)
Public Sub RS_NewRace()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    'As this procedure is triggered automatically and executed only once
    'when opening the workbook comment in the "Stop" command in the
    'following line for debugging purposes
    
'    Stop
    
    Call CreateRaceSheet 'Create the worksheet "GALOPPSIM"
    Call RS_StartScreen 'Draw the start screen
    Call RS_AddControls 'Add menu buttons and a dropdown for the race selection

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RS_NewRace()")
    Call basAuxiliary.CodeCrash
End Sub

'Simulation race for Machine Learning without graphical representation
Public Sub ML_NewRace(lngSimCase As Long, varSave As Variant)
    Dim timestamp As Date 'For saving the current time stamp
    
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Call GetRaceData(True) 'Get the race data from the worksheet with the selected race
    Call AssignBasicValues 'Get basic values
    Call GetHorseData(True) 'Get data about the horses according to the selected race
    Call CheckSpaceRace 'Check whether a space race is chosen
    Call RunRace(True, False) 'Race start
    Call RankNotFinished 'Find the horses that did not finish
    Call TransmitPlacements 'Write the placements to the main array
    
    'Save to hard disk
    If frmMachineLearning.chkExport Then
        Dim intFileNr As Integer 'Output channel
        intFileNr = FreeFile 'Assign the next free number
        
        timestamp = Now
        
        Open varSave For Append As #intFileNr
        
        For i = 1 To UBound(g_arr_varHorses) 'Loop through the horses
            Print #intFileNr, "ID-" & lngSimCase & "-" & timestamp & ";" _
                    & objRace.RACE_ID & ";" _
                    & objRace.PARTICIPANTS & ";" _
                    & objRace.METRES & ";" _
                    & objRace.NUMBER_ENROLLED & ";" _
                    & objRace.NUMBER_STARTING & ";" _
                    & objRace.LANES_FIX_OR_RANDOM & ";" _
                    & objRace.RACE_TYPE & ";" _
                    & objOption.REFUSE_RUN & ";" _
                    & objOption.REFUSAL_RATE & ";" _
                    & objOption.SLIPSTREAM_IMPACT & ";" _
                    & g_arr_varHorses(i, 1) & " (#" & g_arr_varHorses(i, 11) & ")" & ";" _
                    & g_arr_varHorses(i, 0) & ";" _
                    & g_arr_varHorses(i, 12) & ";" _
                    & g_arr_varHorses(i, 15) & ";" _
                    & g_arr_varHorses(i, 5) & ";" _
                    & g_arr_varHorses(i, 6) & ";" _
                    & g_arr_varHorses(i, 25)
        Next i
        
        Close #intFileNr
    End If
    
    'Write to debug window
    If frmMachineLearning.chkDebug Then
        Debug.Print "ID:                " & "ID-" & lngSimCase & "-" & timestamp
        Debug.Print "PARTICIPANTS:      " & objRace.PARTICIPANTS
        Debug.Print "METRES:            " & objRace.METRES
        Debug.Print "ENROLLED:          " & objRace.NUMBER_ENROLLED
        Debug.Print "STARTERS:          " & objRace.NUMBER_STARTING
        Debug.Print "STARTING_ORDER:    " & objRace.LANES_FIX_OR_RANDOM
        Debug.Print "RACETYPE:          " & objRace.RACE_TYPE
        Debug.Print "REFUSE_TO_RUN:     " & objOption.REFUSE_RUN
        Debug.Print "REFUSAL_RATE:      " & objOption.REFUSAL_RATE
        Debug.Print "SLIPSTREAM_IMPACT: " & objOption.SLIPSTREAM_IMPACT
        Debug.Print vbNewLine & "Final ranking list:"
        For i = 1 To UBound(m_arr_varResults) 'Loop through the results
            Debug.Print m_arr_varResults(i, 1) _
                        & ". #" & m_arr_varResults(i, 2) _
                        & " " & m_arr_varResults(i, 3)
        Next i
        Debug.Print
        For i = 1 To UBound(g_arr_varHorses) 'Loop through the horses
            Debug.Print "NAME: " & g_arr_varHorses(i, 1) & " (#" & g_arr_varHorses(i, 11) & ")"
            Debug.Print "    STATUS:       " & g_arr_varHorses(i, 0)
            Debug.Print "    PLACEMENT:    " & g_arr_varHorses(i, 12)
            Debug.Print "    BOX_NR:       " & g_arr_varHorses(i, 15)
            Debug.Print "    BASIC_SPEED:  " & g_arr_varHorses(i, 5)
            Debug.Print "    FORM_FACTOR:  " & g_arr_varHorses(i, 6)
            Debug.Print "    TACTICS:      " & g_arr_varHorses(i, 25)
        Next i
    End If

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "ML_NewRace()")
    Call basAuxiliary.CodeCrash
End Sub

'Get the selected race from the dropdown menu
Private Sub GetRace()
    If g_strPlayMode = "RS" Then
        objRace.SELECTED = g_oleComboRaces.Object.Value 'Get the race which is visible in the dropdown
    Else
        '(currently not used in the AI edition)
    End If
End Sub

'Procedure for painting a horse
Private Sub PaintHorse(ByVal row As Integer, tail As Integer, _
                        colour As Variant)
    Dim horseColour As Long
    
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Call basAuxiliary.ActivateRaceSheet 'Ensure that the
                            '"GALOPPSIM" worksheet is activated
    
    If IsArray(colour) Then 'Multicoloured horse
        Dim count As Integer, countInitial As Integer
        Dim currentColour As Long
        count = 0 'Reset the segment counter
        
        Do While count < 8 'Loop through the 8 segments
                                                'of a horse
            countInitial = count 'Set the counter to the
                                            'current segment
            If IsNumeric(colour(count)) Then
                currentColour = colour(count)
                        'Colour code of the current segment
            Else 'No valid colour has been provided
                currentColour = 3291720 'Use brown instead
            End If
            
            Do 'Check whether adjoining segments have
                                            'the same colour
                If count = 7 Then Exit Do
                If currentColour = colour(count + 1) Then _
                    count = count + 1 Else Exit Do
            Loop

            Range(Cells(row, tail + countInitial), _
                Cells(row, tail + count)).Interior.Color _
                    = currentColour

            count = count + 1 'Next segment
        Loop
        
    Else 'Monochrome horse (no array submitted)
        If IsNumeric(colour) Then
            horseColour = colour
        Else 'No valid colour has been provided
            horseColour = 3291720 'Use brown instead
        End If
        
        Range(Cells(row, tail), _
            Cells(row, tail + 7)).Interior.Color _
                = horseColour 'Paint the horse as a whole
    End If
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, _
        Application.VBE.ActiveCodePane.CodeModule, "PaintHorse()")
    Call basAuxiliary.CodeCrash
End Sub

'Check whether placing bets is allowed for this race
Private Sub CheckBettingAllowed()
    If objOption.BET_MODE = True And objRace.BETTING_ALLOWED = "N" Then
        Call ShowInfoPopup(objRace.RACE_NAME & " " & objRace.RACE_YEAR, _
        GetText(g_arr_Text, "BET050"), _
        False, vbModal, 22)
    End If
End Sub

'Check whether a space race is chosen
Private Sub CheckSpaceRace(Optional test As Boolean)
    'Check whether the Race ID begins with "SPACE"
    If Len(objRace.RACE_ID) > 5 Then Exit Sub
    If left(objRace.RACE_ID, 5) <> "SPACE" Then Exit Sub
    
    'Set gravity values for space races
    Select Case objRace.TRACK_SURFACE
        Case "MOON" 'Moon
            objSpeed.SPEED_LOOP_LOW = -1600
            objSpeed.SPEED_LOOP_HIGH = 2200
        Case "MARS" 'Mars
            objSpeed.SPEED_LOOP_LOW = -1200
            objSpeed.SPEED_LOOP_HIGH = 1800
        Case "JUPITER" 'Jupiter
            objSpeed.SPEED_LOOP_LOW = 0
            objSpeed.SPEED_LOOP_HIGH = 500
        Case "PLUTO" 'Pluto
            objSpeed.SPEED_LOOP_LOW = -5000
            objSpeed.SPEED_LOOP_HIGH = 10000
        Case "SATURN" 'Saturn
            objSpeed.SPEED_LOOP_LOW = -2800
            objSpeed.SPEED_LOOP_HIGH = 3200
    End Select

    Dim strSpaceInfo As String
    
    'No tactics in space
    If objOption.TACTICS Then _
        strSpaceInfo = GetText(g_arr_Text, "SPACEINFO001") & vbNewLine
    'No slipstream in space
    If objOption.SLIPSTREAM_IMPACT > 0 Then _
        strSpaceInfo = strSpaceInfo & GetText(g_arr_Text, "SPACEINFO002") & vbNewLine
    'No start refusal in space
    If objOption.REFUSE_RUN Then _
        strSpaceInfo = strSpaceInfo & GetText(g_arr_Text, "SPACEINFO003")
    
    'Show a pop-up if at least one of the forbidden options is selected
    If Not test And (objOption.TACTICS Or objOption.SLIPSTREAM_IMPACT > 0 Or objOption.REFUSE_RUN) Then _
        Call ShowInfoPopup(objRace.RACE_NAME, objRace.RACE_NAME & vbNewLine & vbNewLine & strSpaceInfo, False, vbModal, 22)
    
End Sub

'Add controls to the GALOPPSIM worksheet in the RS edition
Public Sub RS_AddControls()

    'As this procedure is triggered automatically and executed only once
    'when opening the workbook comment in the "Stop" command in the
    'following line for debugging purposes
'    Stop

    On Error GoTo ERRORHANDLING 'In case an error occurs

    Dim captionStart As String 'Text for the start button dependent on the possiblity to place bets
    captionStart = basAuxiliary.getCaptionStartBtn(objOption.BET_MODE)
    
    'Prepare a collection for the menu buttons
        Set g_colRSbuttons = New Collection
        
    'Add buttons to the menu area
        '"name(ID)", left, top, width, height, font-size, font:bold, _
            background-colour (hex), foreground-colour (hex), caption with the initial text
        Call RS_AddButton("raceoptions", 15, 40, 81, 49, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN001"))
        Call RS_AddButton("exceloptions", 99, 40, 81, 49, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN002"))
        Call RS_AddButton("startrace", 196, 40, 81, 49, 11, True, 52377, 0, GetText(g_arr_Text, captionStart)) 'Green button
        Call RS_AddButton("finishphoto", 280, 40, 81, 49, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN004"))
        Call RS_AddButton("results", 364, 40, 81, 24, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN005"))
        Call RS_AddButton("winner", 364, 65, 81, 24, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN006"))
        Call RS_AddButton("bets", 448, 40, 81, 49, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN007"))
        Call RS_AddButton("saverace", 629, 40, 81, 49, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN023"))
        Call RS_AddButton("loadrace", 713, 40, 81, 49, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN024"))
        Call RS_AddButton("replay", 545, 40, 81, 49, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN025"))
        Call RS_AddButton("language", 810, 40, 81, 24, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "LANGUAGE001"))
        Call RS_AddButton("info", 894, 40, 81, 24, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN009"))
        Call RS_AddButton("warning", 894, 65, 81, 24, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN010"))
        Call RS_AddButton("movie2017", 978, 40, 81, 24, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN011"))
        Call RS_AddButton("colours", 810, 65, 81, 24, 11, False, &HFFFFFF, 0, GetText(g_arr_Text, "BTN021"))
        Call RS_AddButton("developer", 1075, 40, 81, 49, 11, False, &H80000010, &HFFFFFF, GetText(g_arr_Text, "BTN022"))
    
    'Add a combobox for the race selection to the menu area
        Call RS_AddComboboxRaces("CBraces", 196, 15, 598, 20) '"name(ID)", left, top, width, height
    
    'Deactivate some buttons as they have no function before the race
        Call RS_InactivateCommandButtons

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RS_AddControls()")
    Call basAuxiliary.CodeCrash
End Sub

'Set some buttons inactive as they have no function before the race
Public Sub RS_InactivateCommandButtons()
    Call basAuxiliary.ActivateRaceSheet
        With g_wksRace
            .OLEObjects("finishphoto").Object.Enabled = False
            .OLEObjects("results").Object.Enabled = False
            .OLEObjects("winner").Object.Enabled = False
            .OLEObjects("bets").Object.Enabled = False
            .OLEObjects("saverace").Object.Enabled = False
            If Not objRace.STARTED Then .OLEObjects("replay").Object.Enabled = False
        End With
End Sub

'Click on a menu button on the GALOPPSIM worksheet in the Run Simple Edition
Public Sub RS_ExecuteClick(name As String)
    On Error GoTo ERRORHANDLING 'In case an error occurs

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Select Case name 'Determine which button has been clicked
        Case "startrace"
            'Leave the current race and start a new one?
            If objRace.STARTED Then
                
'                'A MessageBox cannot handle unicode, so for example Cyrillic characters are displayed as question marks
'                'Comment in the following lines and switch to Bulgarian language to experience the effect...
'                If MsgBox((GetText(g_arr_Text, "RACE003") & ": " & g_wksRace.OLEObjects("CBraces").Object.text), _
'                    vbOKCancel, g_c_tool) = vbCancel Then Exit Sub
                        
                'Show a UserForm instead that is designed in the style close to a MessageBox
                Call ShowMessagePopup(g_c_tool, GetText(g_arr_Text, "RACE003") & ": " & g_wksRace.OLEObjects("CBraces").Object.text, _
                   enumButton.CancelOK, vbModal)
                
                'Evaluate the return value
'                    'Comment in the following lines for understanding the comparison of different data types
'                    Debug.Print g_enumButton & " " & TypeName(g_enumButton) 'Type of the variable: String
'                    Debug.Print enumButton.Cancel & " " & TypeName(enumButton.Cancel) 'Type of the Enumeration: Long
'                    Debug.Print g_enumButton = enumButton.Cancel 'Even True if the String value ("3") is being compared with the Long value (3)!
                If g_enumButton = enumButton.Cancel Then Exit Sub 'No new race if "Cancel" has been clicked
                
                Call ShowNewRaceScreen("NEWRACE2") 'Show the GALOPPSIM title screen
            End If
            
            Call GetRace 'Get the selected race from the dropdown menu
            Call AssignRaceSheet(False)
            Call GetRaceData(False) 'Read the data of the selected race from the worksheet
            
            Call RS_InactivateCommandButtons 'Inactivate some buttons (RS edition only)
            Call NewRace 'Main procedure for starting a new race
        
        Case "finishphoto"
            Call ShowFinishPhoto
        Case "results"
            Call ShowRankingList
        Case "winner"
            Call ShowWinnerPhoto(False, True)
        Case "bets"
            Call ShowBets
        Case "raceoptions"
            Call GetRace 'Get the selected race from the dropdown menu
            Call GetAnimalGrammar 'Get grammar components for compiling texts with different race participants
            frmOptionsRace.show (vbModal) 'Display the pop-up
        Case "exceloptions"
            frmOptionsExcel.show (vbModal)
        Case "language"
            frmRS_languages.show (vbModal)
            If objRace.SELECTED = "" Then
                objRace.SELECTED = g_oleComboRaces.Object.Value
                Call AssignRaceSheet(False)
                Call GetRaceData(False)
                Call GetAnimalGrammar
            End If
            Call ChangeLanguage
        Case "info"
            Call ShowInfo
        Case "warning"
            Call ShowWarning
        Case "movie2017"
            Call GaloppSimMovie2017
        Case "colours"
            Call ColourModeSelection
        Case "developer" 'Call the Super Developer Tools
            frmSuperDev.show (vbModeless)
        Case "saverace"
            Call SaveRaceForReplay(False)
        Case "loadrace"
            Call LoadRaceForReplay
        Case "replay"
            #If Debugging Then
                Debug.Print "-----> Replay started"
            #End If
            Call RaceReplay
            #If Debugging Then
                Debug.Print "Replay ended <-----"
            #End If
    End Select
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RS_ExecuteClick()")
    Call basAuxiliary.CodeCrash
End Sub

'Show the GALOPPSIM title screen
Public Sub ShowNewRaceScreen(pic As String)
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Call basAuxiliary.ActivateRaceSheet
    
    'Delete speed chart
    Dim sh As Shape
    For Each sh In g_wksRace.Shapes
       If sh.name = "SpeedChart" Then sh.Delete
    Next
    
    'Adjust the column width and row height according to the screen size
    With g_wksRace.UsedRange
        .ColumnWidth = ZoomLevelPictures()(0)
        .RowHeight = ZoomLevelPictures()(1)
        .Clear
    End With
    
    If g_strPlayMode = "RS" Then Call RS_MenuAreaHide 'Hide the controls on the worksheet (RS edition)
    If g_strPlayMode = "AI" Then Call AI_ExcelModeStart
    ActiveWindow.ScrollColumn = 1 'Scroll to the left (column A)
    Call PaintPicture(g_wksPIC, g_wksRace, pic, 100, 40, 1, 1) 'Paint the GALOPPSIM title picture
    Call CursorAway 'Place the cursor far away (in the upper right corner of the screen)
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "ShowNewRaceScreen()")
    Call basAuxiliary.CodeCrash
End Sub

'Add a new button to the worksheet in the RS edition
Private Sub RS_AddButton(n As String, l As Integer, t As Integer, w As Integer, _
                            h As Integer, fs As Integer, fb As Boolean, bc As Long, _
                            fc As Long, c As String)
            'n = name(ID), l = left, t = top, w = width, h = height, fs = font-size,
            'fb = font:bold, bc = background-colour (hex), fc = foreground-colour,
            'c =caption with the initial text
    
    'As this procedure is triggered automatically and executed only once for each button
    'when opening the workbook comment in the "Stop" command in the
    'following line for debugging purposes
'    Stop
                            
    Dim oleRSbutton As OLEObject
    Dim objRSbutton As clsRSbutton

    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    'Add a button
    Set oleRSbutton = g_wksRace.OLEObjects.Add(classtype:="Forms.CommandButton.1", _
    left:=l, top:=t, width:=w, Height:=h)
    
    'Assign properties to the OLE button object using nested "With" commands
    With oleRSbutton
        .name = n 'Button ID [Full command: oleRSbutton.name = n]
        With .Object
            .caption = c '[Full command: oleRSbutton.Object.Caption = c]
            .Font.size = fs '[Full command: oleRSbutton.Object.Font.Size = fs]
            .Font.Bold = fb
            .BackColor = bc
            .ForeColor = fc
            .WordWrap = True
            .TakeFocusOnClick = False
        End With
        .Placement = xlFreeFloating '[Full command: oleRSbutton.Placement = xlFreeFloating]
        .Visible = True
    End With

    'Prepare a new button collection
    Set objRSbutton = New clsRSbutton 'Create a new instance for a button object with event handling
    Set objRSbutton.RSButtonObject = oleRSbutton.Object 'Assign the OLE button
    objRSbutton.RSbtnID = n 'Assign the name of the OLE button which serves as an ID

    'Add the button to the button collection
    g_colRSbuttons.Add objRSbutton
 
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RS_AddButton()")
    Call basAuxiliary.CodeCrash
End Sub

'Add a new combobox with all installed races on the worksheet in the RS edition
Private Sub RS_AddComboboxRaces(n As String, l As Integer, t As Integer, w As Integer, h As Integer)
    On Error GoTo ERRORHANDLING 'In case an error occurs

    'As this procedure is triggered automatically and executed only once
    'when opening the workbook comment in the "Stop" command in the
    'following line for debugging purposes
'    Stop
    
    Dim wksCheck As Worksheet

    Set g_colRacesInstalled = Nothing
    Set g_colRacesInstalled = New Collection

    Set g_oleComboRaces = g_wksRace.OLEObjects.Add(classtype:="Forms.ComboBox.1", _
        left:=l, top:=t, width:=w, Height:=h)

    With g_oleComboRaces
        .name = n 'ID
        .Placement = xlFreeFloating
        .Object.ColumnCount = 2 'Column 0: Name of the Workwheet (e.g. race_DEUTSCHESDERBY17) // Column 1: Visible name
        .Object.ColumnWidths = "0 Pt" 'Width of the first column (0): 0 pixels --> hidden
        .Object.Style = fmStyleDropDownList 'Allow only values from the item list, no free entries
        .Visible = True
    End With
    
    'Populate the dropdown with all installed races (all worksheets beginning with "race_")
    For Each wksCheck In ThisWorkbook.Worksheets
        If left(wksCheck.name, 5) = "race_" Then
            If wksCheck.Cells(basAuxiliary.GetRow(wksCheck, "STATUS"), basAuxiliary.GetColumn(wksCheck, "RACE DATA VALUE")).Value = "released" Then
                g_colRacesInstalled.Add wksCheck.name
                With g_oleComboRaces.Object
                .AddItem
                .List(.ListCount - 1, 0) = wksCheck.name
                .List(.ListCount - 1, 1) = wksCheck.Cells(basAuxiliary.GetRow(wksCheck, "RACE NAME"), 2).Value & " " & _
                                        wksCheck.Cells(basAuxiliary.GetRow(wksCheck, "YEAR"), 2).Value & " (" & _
                                        wksCheck.Cells(basAuxiliary.GetRow(wksCheck, "DISTANCE METRES"), 2).Value & "m) - " & wksCheck.Cells(basAuxiliary.GetRow(wksCheck, "TRACK LOCATION"), 2).Value
                End With
            End If
        End If
    Next wksCheck

    'Set default race
    g_oleComboRaces.Object.ListIndex = 0 'take the first race

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RS_AddComboboxRaces()")
    Call basAuxiliary.CodeCrash
End Sub

'Hide all controls on the worksheet in the RS edition
Private Sub RS_MenuAreaHide()
    Dim oleObj As OLEObject
    
    Call basAuxiliary.ActivateRaceSheet
    
    'Loop through all control objects and hide them one by one
    For Each oleObj In g_wksRace.OLEObjects
        #If Debugging Then
            Debug.Print "Hide OLEObject: " & oleObj.name
        #End If
        oleObj.Visible = False
    Next oleObj
    
    'Hide the top rows which are used for the control objects
    Range(rows(2), rows(objBasicData.TOP_ROWS)).Hidden = True
End Sub

'Show all controls on the worksheet in the RS edition
Public Sub RS_MenuAreaShow()
    Dim oleObj As OLEObject
    
    Call basAuxiliary.ActivateRaceSheet

    'Show the top rows which are used for the control objects
    Range(rows(2), rows(objBasicData.TOP_ROWS)).Hidden = False

    'Loop through all control objects and show them one by one
    For Each oleObj In g_wksRace.OLEObjects
        #If Debugging Then
            Debug.Print "Show OLEObject: " & oleObj.name
        #End If
        oleObj.Visible = True
    Next oleObj

End Sub

'Start screen for the RS edition
Public Sub RS_StartScreen()

    'As this procedure is triggered automatically
    'when opening the workbook comment in the "Stop" command in the
    'following line for debugging purposes
'    Stop

    Cells.Clear 'Clear the whole worksheet
    
    'Paint the title picture
    Call PaintPicture(g_wksPIC, g_wksRace, "RUNSIMPLE4", 100, 40, 1, 1)
    'Formattings for the area with the picture
    With Range(Columns(1), Columns(100)) 'Column width and row height dependent on the window size
        .ColumnWidth = ZoomLevelPictures()(0)
        .RowHeight = ZoomLevelPictures()(1)
        .rows(1).RowHeight = 8 'Height of the top row
        With .rows(2) 'Formattings for the second row
            .Font.name = "Arial Black" 'For getting the font in the dropdown with the races bold
            .EntireRow.AUTOFIT
        End With
    End With
    'Write the title in cell B2
    Cells(2, 2).Value = g_c_tool & " (v" & g_c_version & ")"
    'Place the cursor far away (in the upper right corner of the screen)
    Call CursorAway
End Sub

'Create the worksheet "GALOPPSIM"
Private Sub CreateRaceSheet()
    On Error GoTo ERRORHANDLING 'In case an error occurs

    'Check whether the worksheet already exists
    For Each m_wksCheck In ActiveWorkbook.Worksheets
        If m_wksCheck.name = "GALOPPSIM" Then
            Application.DisplayAlerts = False 'Suppress the warning message
            m_wksCheck.Delete 'Delete the worksheet
            Application.DisplayAlerts = True 'Re-activate warning messages
        End If
    Next m_wksCheck
    'Create a new worksheet
        Set g_wksRace = ActiveWorkbook.Worksheets.Add(Before:=Sheets(1))
        With g_wksRace
            .name = "GALOPPSIM"
            .Activate
        End With

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "CreateRaceSheet()")
    Call basAuxiliary.CodeCrash
End Sub

'Assign basic values
Private Sub AssignBasicValues(Optional test As Boolean)
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    'Font and cell sizes
    m_intTrackCellHeight = 12
    m_dblTrackCellWidth = 0.5
    m_intFontSize = 10
    objBasicData.RANK_WIDTH = 6
    m_intAdvertisingHeight = 6
        
'Version 152.00: At runtime auslesen per GetColumn und so ToDo
    If left(objRace.RACE_ID, 5) = "SPACE" Then
        objSpeed.SPEED_BASIC_LOW = 100
        objSpeed.SPEED_BASIC_HIGH = 100
        objSpeed.SPEED_COND_LOW = 100
        objSpeed.SPEED_COND_HIGH = 100
    Else
        With g_wksTEC
            'Range of the basic speed in case it is not fixed for a horse
            objSpeed.SPEED_BASIC_LOW = .Range("A3").Value 'Standard value: 1480
            objSpeed.SPEED_BASIC_HIGH = .Range("A2").Value 'Standard value: 1520
            'Range of the daily form of the horses
            objSpeed.SPEED_COND_LOW = .Range("B3").Value 'Standard value: 1490
            objSpeed.SPEED_COND_HIGH = .Range("B2").Value 'Standard value: 1510
            'Range of the randomly assigned speed per step
            objSpeed.SPEED_LOOP_LOW = .Range("C3").Value 'Standard value: 0
            objSpeed.SPEED_LOOP_HIGH = .Range("C2").Value 'Standard value: 3000
            'Range of each phase if racing tactics are active
            objSpeed.SPEED_TACTICS_LOW = .Range("D4").Value 'Standard value: 1200
            objSpeed.SPEED_TACTICS_MEDIUM = .Range("D3").Value '1500
            objSpeed.SPEED_TACTICS_HIGH = .Range("D2").Value 'Standard value: 1800
        End With
    End If
    
    'Columns left of the starting line (minimum 7, default value 11)
    objBasicData.LEFT_COLS = 12
    'Columns behind the finish line (minimum 5)
    objBasicData.AFTER_FIN_COLS = 5
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "AssignBasicValues()")
    Call basAuxiliary.CodeCrash
End Sub

'Assign the worksheet with the selected race
Private Sub AssignRaceSheet(replay As Boolean)
    If replay Then
        Dim wksFind As Worksheet
        For Each wksFind In ThisWorkbook.Worksheets
            If left(wksFind.name, 5) = "race_" Then
                If wksFind.Cells(basAuxiliary.GetRow(wksFind, "RACE ID"), _
                    basAuxiliary.GetColumn(wksFind, "RACE DATA VALUE")).Value _
                    = g_arr_varReplay_RaceData(3, 1) Then
                        Set g_wksRaceData = wksFind
                End If
            End If
        Next wksFind
    Else
        Set g_wksRaceData = ThisWorkbook.Worksheets(objRace.SELECTED)
    End If
End Sub

'Read the data of the selected race from the worksheet
Private Sub GetRaceData(ml As Boolean)
    On Error GoTo ERRORHANDLING 'In case an error occurs

    'Read the full race data from the worksheet
    Dim col As Integer 'Column to get the data from
    col = basAuxiliary.GetColumn(g_wksRaceData, "RACE DATA VALUE")
    With g_wksRaceData
        objRace.RACE_ID = .Cells(basAuxiliary.GetRow(g_wksRaceData, "RACE ID"), col).Value 'Unique race ID
        objRace.REAL_RACE = .Cells(basAuxiliary.GetRow(g_wksRaceData, "REAL RACE"), col).Value 'Real race (yes or no)
        objRace.PARTICIPANTS = .Cells(basAuxiliary.GetRow(g_wksRaceData, "PARTICIPANTS"), col).Value 'Race participants (HORSE/PIG/DONKEY/DOG/UNICORN)
        objRace.RACE_NAME = .Cells(basAuxiliary.GetRow(g_wksRaceData, "RACE NAME"), col).Value 'Race name
        objRace.RACE_YEAR = .Cells(basAuxiliary.GetRow(g_wksRaceData, "YEAR"), col).Value 'Year of the race
        objRace.TRACK_LOCATION = .Cells(basAuxiliary.GetRow(g_wksRaceData, "TRACK LOCATION"), col).Value 'Track location
        objRace.COUNTRY_CODE = .Cells(basAuxiliary.GetRow(g_wksRaceData, "COUNTRY"), col).Value 'Country code
        objRace.COUNTRY_NAME = GetCountryName(.Cells(basAuxiliary.GetRow(g_wksRaceData, "COUNTRY"), col), objOption.language) 'Country name
        objRace.TRACK_NAME = .Cells(basAuxiliary.GetRow(g_wksRaceData, "TRACK NAME"), col).Value 'Track name
        objRace.TRACK_COLOUR = .Cells(basAuxiliary.GetRow(g_wksRaceData, "TRACK COLOUR"), col).Value 'Track colour
        objRace.TRACK_SURFACE = .Cells(basAuxiliary.GetRow(g_wksRaceData, "TRACK SURFACE"), col).Value 'Track surface
        objRace.RACE_TYPE = .Cells(basAuxiliary.GetRow(g_wksRaceData, "RACE TYPE"), col).Value 'Race type
        objRace.METRES = .Cells(basAuxiliary.GetRow(g_wksRaceData, "DISTANCE METRES"), col).Value 'Race distance
        objRace.STARTING_GATE = .Cells(basAuxiliary.GetRow(g_wksRaceData, "STARTING GATE"), col).Value 'Starting gate (yes or no)
        objRace.LANES_FIX_OR_RANDOM = .Cells(basAuxiliary.GetRow(g_wksRaceData, "LANES FIX OR RANDOM"), col).Value 'Lanes fix or random
        objRace.ADVERTISING = .Cells(basAuxiliary.GetRow(g_wksRaceData, "ADVERTISING"), col).Value 'Advertising (yes or no)
        objRace.BETTING_ALLOWED = .Cells(basAuxiliary.GetRow(g_wksRaceData, "BETTING ALLOWED"), col).Value 'Betting allowed (yes or no)
        objRace.NUMBER_ENROLLED = .Cells(.rows.count, GetColumn(g_wksRaceData, "STATUS")).End(xlUp).row - 1 'Number of horses enrolled
        objRace.NUMBER_STARTING = Application.WorksheetFunction.CountIf(.Columns(GetColumn(g_wksRaceData, "STATUS")), "START")
        objRace.SPECIAL = .Cells(basAuxiliary.GetRow(g_wksRaceData, "SPECIAL"), col).Value 'For special purposes
        GMP_mode = .Cells(basAuxiliary.GetRow(g_wksRaceData, "GMP_MODE"), col).Value 'For Matjes Grand Prix Races
    End With

    If Not ml Then 'Skip for Machine Learning simulation races
        'Take the colour mode into account
        Select Case g_strColourMode
            Case "POPART"
                If objRace.TRACK_COLOUR > 0 And objRace.TRACK_COLOUR < 16777215 Then _
                objRace.TRACK_COLOUR = PopArtColour(objRace.TRACK_COLOUR)
            Case "LSD"
                If objRace.TRACK_COLOUR > 0 And objRace.TRACK_COLOUR < 16777215 Then _
                objRace.TRACK_COLOUR = PopArtColour(Int((16777215 - 0 + 1) * Rnd + 0))
            Case "SMARTIES"
                If objRace.TRACK_COLOUR > 0 And objRace.TRACK_COLOUR < 16777215 Then _
                objRace.TRACK_COLOUR = Int((16777215 - 0 + 1) * Rnd + 0)
            Case "TV1960"
                objRace.TRACK_COLOUR = GreyToLong(CInt(RGBtoGrey(CLng(objRace.TRACK_COLOUR))))
            Case "DARKMODE"
                objRace.TRACK_COLOUR = 2697513 'Dark grey
            Case "24H"
                objRace.TRACK_COLOUR = DuskDawn(objRace.TRACK_COLOUR, Abs(22 * objOption.DAYLIGHT))
        End Select
    End If
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "GetRaceData()")
    Call basAuxiliary.CodeCrash
End Sub

'Get graphics data (e.g. advertisement, tribunes, food and beverage)
Private Sub GetTrackGraphicsData(strData As String)
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Dim col As Integer 'Column with the advertisement data
    col = basAuxiliary.GetColumn(g_wksRaceData, strData)

        If g_wksRaceData.Cells(rows.count, col).End(xlUp).row > 1 Then 'Only if there is at least 1 entry
            j = g_wksRaceData.Cells(rows.count, col).End(xlUp).row - 1 'Last row with data in the column
            ReDim m_arr_varTrackGraphics(1 To j) 'Location of the data
            For i = 1 To j
                m_arr_varTrackGraphics(i) = g_wksRaceData.Cells(i + 1, col).Value 'Write the value into an array
            Next i
        End If
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "GetTrackGraphicsData()")
    Call basAuxiliary.CodeCrash
End Sub

'Get the name of the country in which the race takes place
Private Function GetCountryName(code As String, language As String) As String
    Dim col As Integer, row As Integer
    
    'Find the column on the worksheet "TEXT"
    col = basAuxiliary.GetColumn(g_wksTEXT, objOption.language)
        
    'Find the row on the worksheet "TEXT"
    row = basAuxiliary.GetRow(g_wksTEXT, code)
        
    'Return the name
    If g_wksTEXT.Cells(row, col).Value = "" Then
        GetCountryName = g_wksTEXT.Cells(row, GetColumn(g_wksTEXT, "EN")).Value 'Take the English name as fallback
    Else
        GetCountryName = g_wksTEXT.Cells(row, col).Value 'Get the name according to the selected language
    End If
End Function

'Read all horse data from the race sheet
Private Sub GetHorseData(Optional ml As Boolean)
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    'Resize the array for the horse data
    ReDim g_arr_varHorses(1 To objRace.NUMBER_ENROLLED, 0 To 31) 'All data of the horses
    
    'In case of a random line-up at the start: Write all starting gates into an array
    If objRace.LANES_FIX_OR_RANDOM = "R" Then
        Dim gateNr As Integer
        Dim inGate As Boolean
        Dim arrGates() As Integer
        ReDim arrGates(1 To objRace.NUMBER_ENROLLED)
        For i = 1 To objRace.NUMBER_ENROLLED
            arrGates(i) = i
        Next i
    End If
    
    'Loop through all horses on the worksheet and get the data
    For i = 1 To objRace.NUMBER_ENROLLED
        Dim arr_colour(0 To 7) 'Array with the colours of the 8 horse segments
        Dim colour As Integer
        Dim same_colour As Boolean
        same_colour = True
        
        With g_wksRaceData
            g_arr_varHorses(i, 11) = .Cells(1 + i, GetColumn(g_wksRaceData, "NR")).Value 'Horse number
            g_arr_varHorses(i, 0) = .Cells(1 + i, GetColumn(g_wksRaceData, "STATUS")).Value 'Status before the race (START, CANCELLED, CORONAPOSITIVE)
            g_arr_varHorses(i, 1) = .Cells(1 + i, GetColumn(g_wksRaceData, "NAME")).Value 'Horse name
        End With
    
        If Not ml Then 'Skip for Machine Learning simulation races
                'Get the horse colours
                If Not IsEmpty(g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "COLOUR 8 (HEAD)"))) Then 'If the colour of the head is not empty
                    For colour = 0 To 7
                        arr_colour(colour) = g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "COLOUR 1 (TAIL)") + colour) 'Segment colour
                        If colour > 0 Then 'If the colour is not empty (or black)
                            If Not g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "COLOUR 1 (TAIL)") + colour) = g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "COLOUR 1 (TAIL)") + colour - 1) Then
                                same_colour = False 'If the colour of this segment differs from that before
                            End If
                        End If
                    Next colour
                    'Assign either a Long value or an Array with 8 fields for the horse colour
                    If same_colour Then g_arr_varHorses(i, 2) = arr_colour(0) Else g_arr_varHorses(i, 2) = arr_colour
                Else 'If no colour of the head is found: Determine a random colour for the whole horse
                    Randomize 'Initialize the random number generator
                    g_arr_varHorses(i, 2) = Int((16777215 - 0 + 1) * Rnd + 0) 'Apply the randomly generated colour for the whole horse
                    'Formula for the generation of a random integer value within a specific range:
                    'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
                    '>> replace upperbound and lowerbound with integer values
                End If
        End If

        'In case of a random line-up at the start: Assign the starting gates
        If objRace.LANES_FIX_OR_RANDOM = "R" Then
            inGate = False 'Horse is not yet assigned to a starting gate
            Do Until inGate = True 'Loop until a starting gate is assigned
                Randomize 'Initialize the random number generator
                gateNr = (Int((objRace.NUMBER_ENROLLED - 1 + 1) * Rnd + 1)) 'Random number
                If arrGates(gateNr) <> 0 Then 'If this gate is empty
                    g_arr_varHorses(i, 15) = gateNr 'Starting gate
                    arrGates(gateNr) = 0 'Mark the gate as occupied
                    inGate = True 'Horse is assigned to a gate
                End If
            Loop
        Else 'If the lanes are fix
            g_arr_varHorses(i, 15) = g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "LANE")).Value 'Read the starting gate from the worksheet
        End If
        
        g_arr_varHorses(i, 3) = objBasicData.TOP_ROWS + 5 + 2 * g_arr_varHorses(i, 15) 'Row number on which the horse will run
        g_arr_varHorses(i, 4) = objBasicData.LEFT_COLS + 12 'Starting position (column number)
        g_arr_varHorses(i, 9) = 0 'Exact position for internal calculation (0 to [race distance in metres * 100])
        
        'Get the basic speed
        If Not IsEmpty(g_wksRaceData.Cells(1 + i, 16)) Then 'If a value is found on the race sheet
            g_arr_varHorses(i, 5) = g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "SPEED")).Value 'Fixed basic speed
        Else 'If no value is found: generate it by random within a determined range
            Randomize
            g_arr_varHorses(i, 5) = Int((objSpeed.SPEED_BASIC_HIGH - objSpeed.SPEED_BASIC_LOW + 1) * Rnd + objSpeed.SPEED_BASIC_LOW)
        End If
        
        'Determine the daily form of the horse by random
        Randomize
        g_arr_varHorses(i, 6) = Int((objSpeed.SPEED_COND_HIGH - objSpeed.SPEED_COND_LOW + 1) * Rnd + objSpeed.SPEED_COND_LOW)
        
        If Not ml Then 'Skip for Machine Learning simulation races
                'Determine the betting odds
                If Not IsEmpty(g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "ODDS (X:10)"))) Then 'If a value is found on the race sheet
                    g_arr_varHorses(i, 17) = g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "ODDS (X:10)")).Value 'Fixed odds
                Else 'If no value is found: derive it from the basic speed with a complex formula
                    Randomize
                    'Rounded integer value from (((50 + (((number of starters + 2) / 6) * (1523 - basic speed) ^ 2)) / 5) * random value between 0.9 and 1.1)
                    g_arr_varHorses(i, 17) = Round(((50 + (((objRace.NUMBER_ENROLLED + 2) / 6) * (1523 - g_arr_varHorses(i, 5)) ^ 2)) / 5) * (Int((11 - 9 + 1) * Rnd + 9) / 10), 0)
                End If
                    
                'Estimation error for the impression during the warm-up (+/-50 pixels of the bar length)
                Randomize
                g_arr_varHorses(i, 18) = (Int((100 - 0 + 1) * Rnd + 0)) - 50 'Random number between -50 and +50
                'Alternatively:
                'g_arr_varHorses(i, 18) = Int((50 - (-50) + 1) * Rnd + (-50))
        End If

        'Reset the slipstream factor
        g_arr_varHorses(i, 22) = 0
            
        If Not ml Then 'Skip for Machine Learning simulation races
            'Get the picture of the winner
            If g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "PHOTO")).Value <> "" Then 'If a value is found on the race sheet
                g_arr_varHorses(i, 23) = g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "PHOTO")).Value 'Specific picture
            Else 'If no value is found: take the default picture
                g_arr_varHorses(i, 23) = "WINNER_" & objRace.PARTICIPANTS & "_DEFAULT"
            End If
        End If

        'Attribute for different purposes like a special race behaviour
        g_arr_varHorses(i, 24) = g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "SPECIAL")).Value
    
        If Not ml Then 'Skip for Machine Learning simulation races
            'Take the colour mode into account
            Select Case g_strColourMode
                Case "POPART"
                    If IsArray(g_arr_varHorses(i, 2)) Then
                        For j = 0 To 7
                            g_arr_varHorses(i, 2)(j) = PopArtColour(g_arr_varHorses(i, 2)(j))
                            
                            If g_arr_varHorses(i, 2)(j) = objRace.TRACK_COLOUR Then
                                Do
                                    g_arr_varHorses(i, 2)(j) = PopArtColour(Int((16777215 - 0 + 1) * Rnd + 0))
                                Loop Until g_arr_varHorses(i, 2)(j) <> objRace.TRACK_COLOUR
                            End If
                        Next j
                    Else
                        g_arr_varHorses(i, 2) = PopArtColour(g_arr_varHorses(i, 2))
                            
                        If g_arr_varHorses(i, 2) = objRace.TRACK_COLOUR Then
                            Do
                                g_arr_varHorses(i, 2) = PopArtColour(Int((16777215 - 0 + 1) * Rnd + 0))
                            Loop Until g_arr_varHorses(i, 2) <> objRace.TRACK_COLOUR
                        End If
                    End If
                Case "LSD"
                    If IsArray(g_arr_varHorses(i, 2)) Then
                        For j = 0 To 7
                            Do
                                g_arr_varHorses(i, 2)(j) = PopArtColour(Int((16777215 - 0 + 1) * Rnd + 0))
                            Loop Until g_arr_varHorses(i, 2)(j) <> objRace.TRACK_COLOUR
                        Next j
                    Else
                        Do
                            g_arr_varHorses(i, 2) = PopArtColour(Int((16777215 - 0 + 1) * Rnd + 0))
                        Loop Until g_arr_varHorses(i, 2) <> objRace.TRACK_COLOUR
                    End If
                Case "SMARTIES"
                    If IsArray(g_arr_varHorses(i, 2)) Then
                        For j = 0 To 7
                            g_arr_varHorses(i, 2)(j) = Int((16777215 - 0 + 1) * Rnd + 0)
                        Next j
                    Else
                        g_arr_varHorses(i, 2) = Int((16777215 - 0 + 1) * Rnd + 0)
                    End If
                Case "TV1960"
                    If IsArray(g_arr_varHorses(i, 2)) Then
                        For j = 0 To 7
                            g_arr_varHorses(i, 2)(j) = GreyToLong(CInt(RGBtoGrey(CLng(g_arr_varHorses(i, 2)(j)))))
                        Next j
                    Else
                        g_arr_varHorses(i, 2) = GreyToLong(CInt(RGBtoGrey(CLng(g_arr_varHorses(i, 2)))))
                    End If
                Case "DARKMODE"
                    If IsArray(g_arr_varHorses(i, 2)) Then
                        For j = 0 To 7
                            g_arr_varHorses(i, 2)(j) = DarkModeColour(g_arr_varHorses(i, 2)(j))
                            
                            If g_arr_varHorses(i, 2)(j) = objRace.TRACK_COLOUR Then _
                                g_arr_varHorses(i, 2)(j) = DarkModeColour(Int((16777215 - 0 + 1) * Rnd + 0))
                        Next j
                    Else
                        g_arr_varHorses(i, 2) = DarkModeColour(g_arr_varHorses(i, 2))
                            
                        If g_arr_varHorses(i, 2) = objRace.TRACK_COLOUR Then _
                            g_arr_varHorses(i, 2) = DarkModeColour(Int((16777215 - 0 + 1) * Rnd + 0))
                    End If
                Case "24H"
                    If IsArray(g_arr_varHorses(i, 2)) Then
                        For j = 0 To 7
                            g_arr_varHorses(i, 2)(j) = DuskDawn(g_arr_varHorses(i, 2)(j), Abs(22 * objOption.DAYLIGHT))
                        Next j
                    Else
                        g_arr_varHorses(i, 2) = DuskDawn(g_arr_varHorses(i, 2), Abs(22 * objOption.DAYLIGHT))
                    End If
            End Select
        End If

        Next i
            
    'Determine the speed of each horse in each phase of the race
    '(even if is not used in the race when tactics are deactivated)
    For i = 1 To objRace.NUMBER_ENROLLED
        If g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "TACTICS")).Value <> "" Then 'If a value is found on the race sheet
            g_arr_varHorses(i, 25) = g_wksRaceData.Cells(1 + i, GetColumn(g_wksRaceData, "TACTICS")).Value
        Else
            g_arr_varHorses(i, 25) = g_wksTEC.Cells( _
                Int((142 - 2 + 1) * Rnd + 2), _
                GetColumn(g_wksTEC, "TACTICS")).Value
        End If
        
        For j = 1 To 6 'Convert the letters to speed values
            g_arr_varHorses(i, 25 + j) = TacticMapping(Mid(g_arr_varHorses(i, 25), j, 1))
        Next j
    Next i
    
    'Prepare the momentum log
    Dim arrTemp() As Double
    ReDim arrTemp(1 To objOption.MOMENTUM_REFRESHRATE)
    For i = 1 To objRace.NUMBER_ENROLLED
        g_arr_varHorses(i, 19) = arrTemp
    Next i
    
    'Prepare the race speed log
    ReDim arrTemp(1 To 100) 'Initial length
    For i = 1 To objRace.NUMBER_ENROLLED
        g_arr_varHorses(i, 14) = arrTemp
    Next i
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "GetHorseData()")
    Call basAuxiliary.CodeCrash
End Sub

'Calculation of the race favourites
Private Sub CalculateFavourites()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Erase m_dblFavCalc 'Clear the entire array
        
'        'Alternatively: Clear the array fields one by one
'        m_dblFavCalc(1) = 0
'        m_dblFavCalc(2) = 0
'        m_dblFavCalc(3) = 0
'
'        'Alternatively: Clear the array fields by using a loop
'        For i = 1 To 3
'             m_dblFavCalc(i) = 0
'        Next i
        
    'Calculation of the three favourites by summing up the basic speed and the daily form
    For i = 1 To objRace.NUMBER_ENROLLED
        If g_arr_varHorses(i, 0) = "START" Then
            If g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) > m_dblFavCalc(1) Then
                m_dblFavCalc(3) = m_dblFavCalc(2)
                m_dblFavCalc(2) = m_dblFavCalc(1)
                m_dblFavCalc(1) = g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6)
                m_byteFavourite(3) = m_byteFavourite(2)
                m_byteFavourite(2) = m_byteFavourite(1)
                m_byteFavourite(1) = i 'Horse number of the favourite
            ElseIf g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) > m_dblFavCalc(2) Then
                m_dblFavCalc(3) = m_dblFavCalc(2)
                m_dblFavCalc(2) = g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6)
                m_byteFavourite(3) = m_byteFavourite(2)
                m_byteFavourite(2) = i 'Horse number of another favourite
            ElseIf g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) > m_dblFavCalc(3) Then
                m_dblFavCalc(3) = g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6)
                m_byteFavourite(3) = i 'Horse number of another favourite
            End If
        End If
    Next i
    
    'Write the favourites into an array
    g_arr_varHorses(m_byteFavourite(1), 16) = 1
    g_arr_varHorses(m_byteFavourite(2), 16) = 2
    g_arr_varHorses(m_byteFavourite(3), 16) = 3
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "CalculateFavourites()")
    Call basAuxiliary.CodeCrash
End Sub

Private Function TacticMapping(str As String)
    Select Case str
        Case "S":
            TacticMapping = 1200
        Case "M":
            TacticMapping = 1500
        Case "F":
            TacticMapping = 1800
        Case Else 'e.g. in case of special (numerical) tactics
            TacticMapping = 1500
    End Select
End Function

'Reset all Excel settings
Public Sub ResetExcelOptions()
    Call SetExcelOptions(True, True, True, _
                            True, True, True, _
                            True, False, True)
End Sub

'Turn on the GaloppSim TV mode but with the Excel ribbon visible (only in AI edition)
Public Sub ExcelOptionsTVmenu()
    Call SetExcelOptions(False, False, False, _
                             False, False, False, _
                             False, False, True)
End Sub

'Turn on the GaloppSim TV mode
Public Sub ExcelOptionsTVfull()
    Call SetExcelOptions(False, False, False, _
                             False, False, False, _
                             False, True, True)
End Sub

'Execute the change of the Excel settings
Private Sub SetExcelOptions(blnGrid As Boolean, blnHead As Boolean, blnFormula As Boolean, _
                            blnStatus As Boolean, blnVScroll As Boolean, blnHScroll As Boolean, _
                            blnTabs As Boolean, blnFull As Boolean, blnMax As Boolean)

    With Application
        'Since some parameters depend on each other, the order of execution is important
            .DisplayFullScreen = blnFull 'Excel ribbon
            .ActiveWindow.DisplayGridlines = blnGrid 'Gridlines
            .ActiveWindow.DisplayHeadings = blnHead 'Row and column headings
            .DisplayFormulaBar = blnFormula 'Formula bar
            .DisplayStatusBar = blnStatus 'Status bar
            .ActiveWindow.DisplayVerticalScrollBar = blnVScroll 'Vertical scrollbar
            .ActiveWindow.DisplayHorizontalScrollBar = blnHScroll 'Horizontal scrollbar
            .ActiveWindow.DisplayWorkbookTabs = blnTabs 'Workbook tabs
        If blnMax = True Then 'Window size
            .ActiveWindow.WindowState = xlMaximized
        Else
            .ActiveWindow.WindowState = xlNormal
        End If
    End With

End Sub

'Determine the fitting column width and row height for pictures according to the window size
Public Function ZoomLevelPictures() As Variant() 'Return an array with the values for the column width and row height
        Dim dblWindowHeight As Double
        
        dblWindowHeight = Application.ActiveWindow.Height 'Window height
        
        Select Case dblWindowHeight
            Case Is > 1100 'Large window (higher than 1100 pixels)
                ZoomLevelPictures = Array(3, 22) 'Column width, row height
            Case Is > 799 'Medium-sized window
                ZoomLevelPictures = Array(2, 15)
            Case Else 'Small window
                ZoomLevelPictures = Array(1.8, 14)
        End Select
End Function

'Grammar components for compiling texts with different race participants
Public Sub GetAnimalGrammar()
    Dim animal As String 'Race participants (HORSE/PIG/DONKEY/DOG/UNICORN)
    Dim col As Integer

    If objRace.SELECTED = "" Then objRace.SELECTED = g_wksRaceData.name 'In case of an automatic test execution

    With ThisWorkbook 'Read the type of participants from the worksheet
        animal = .Worksheets(objRace.SELECTED).Cells(basAuxiliary.GetRow(.Worksheets(objRace.SELECTED), "PARTICIPANTS"), 2).Value
    End With
    
    'Get the column with the language
    col = basAuxiliary.GetColumn(g_wksTEXT, objOption.language)
    
    'Get the text component
    For i = 1 To g_wksTEXT.Cells(rows.count, 1).End(xlUp).row
        If g_wksTEXT.Cells(i, 1).Value = animal Then
            g_arr_Grammar(g_wksTEXT.Cells(i, 2).Value) = g_wksTEXT.Cells(i, col).Value
        End If
    Next i
    
End Sub

'Draw the race track when starting a new race
Private Sub DrawRaceTrack(replay As Boolean)

    On Error GoTo ERRORHANDLING 'In case an error occurs

    Dim intPicColumn As Integer 'Column that contains graphics data

    'Variables for overriding settings
    Dim intNumberSpectators As Integer

    Call basAuxiliary.ActivateRaceSheet
    
    'Deactivate screen updating
    Application.ScreenUpdating = False
        
    'Freeze columns A-M if one of those checkboxes is activated, otherwise unfreeze
    If objOption.NAMES_LEFT Or objOption.COLOURS_LEFT Or objOption.MOMENTUM_BARS _
        Or objOption.MOMENTUM_ICONS Or objOption.TACTICS_REVEAL_TAC Or objOption.TACTICS_REVEAL_CURR _
        Or objOption.HIGHLIGHT_FAV Or (objOption.RACE_INFO And objOption.RACE_INFO_WKS) _
        Or (objOption.FOCUSED_RUN <> enumCamera.standard And objOption.HIGHLIGHT_FOC) Then
            Call basAuxiliary.Freeze(13, 0, True) 'Freeze
    Else
            Call basAuxiliary.Freeze(0, 0, False) 'Unfreeze
    End If

    'Formatting: Row height of different sections
    Range(rows(1 + objBasicData.TOP_ROWS), rows(5 + objBasicData.TOP_ROWS)).EntireRow.RowHeight = 15 'Above the race track
    Range(rows(objRace.NUMBER_ENROLLED * 2 + 6 + 1 + objBasicData.TOP_ROWS), _
        rows(objRace.NUMBER_ENROLLED * 2 + 52 + objBasicData.TOP_ROWS)).EntireRow.RowHeight = 15 'Below the race track
    rows(objRace.NUMBER_ENROLLED * 2 + 20 + objBasicData.TOP_ROWS).RowHeight = 20 'Headline of the ranking list

    'Display race data on the top
    With Cells(2 + objBasicData.TOP_ROWS, 14) 'Race name, year, track and location
        .Font.name = "Arial Black"
        .Font.Color = objOption.COL_TEXT

        If replay Then
            .Value = "[" & UCase(GetText(g_arr_Text, "RACE035")) & "] " & objRace.RACE_NAME & " " & objRace.RACE_YEAR & " - " & objRace.TRACK_NAME & ", " & objRace.TRACK_LOCATION _
                & " (" & objRace.COUNTRY_NAME & ")"
            .Characters(Start:=1, Length:=(Len(GetText(g_arr_Text, "RACE035")) + 2)).Font.Color = vbRed
        Else
            .Value = objRace.RACE_NAME & " " & objRace.RACE_YEAR & " - " & objRace.TRACK_NAME & ", " & objRace.TRACK_LOCATION _
                & " (" & objRace.COUNTRY_NAME & ")"
        End If
    End With
    With Cells(3 + objBasicData.TOP_ROWS, 14) 'Race and track type
        .Font.name = "Arial"
        .Font.Color = objOption.COL_TEXT
        .Font.Bold = True
        .Value = objRace.RACE_TYPE_TEXT & " " & GetText(g_arr_Text, "RACE007") & " " & _
            objRace.METRES & " " & GetText(g_arr_Text, "RACE009") & " - " & objRace.TRACK_SURFACE_TEXT
    End With
    
    'Formatting: Columns on the left of the starting grid
    Columns(1).ColumnWidth = 6
    Range(Columns(2), Columns(9)).ColumnWidth = m_dblTrackCellWidth 'Horse colours
    Columns(10).ColumnWidth = 3 'Current race section speed
    Columns(11).ColumnWidth = 6 'Horse numbers
    Columns(12).ColumnWidth = 22
    Range(Columns(14), Columns(objBasicData.LEFT_COLS + 12)).ColumnWidth = m_dblTrackCellWidth 'Starting area
    Columns(objBasicData.LEFT_COLS + 4).ColumnWidth = 6 'Starting gate numbers
    Columns(13).ColumnWidth = 3 'Momentum speed bars
    
    'Formatting: Background colour (behind the track)
    Cells.Interior.Color = objOption.COL_BACK
        
    'Formatting: Race track to run (1 metre = 1 column)
    Range(Columns(objBasicData.LEFT_COLS + 13), Columns(objRace.METRES + objBasicData.LEFT_COLS + 13 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))).ColumnWidth = m_dblTrackCellWidth 'Column width
    'Row height: Alternating higher and lower
    For i = (6 + objBasicData.TOP_ROWS) To (objRace.NUMBER_ENROLLED * 2 + 6 + objBasicData.TOP_ROWS)
        If objRace.SPECIAL = "RACIALDISTANCING" Then 'Racial distancing(Coronavirus rules)
            rows(i).RowHeight = m_intTrackCellHeight * (2 - (i - (6 + objBasicData.TOP_ROWS)) Mod 2)
        Else 'Standard distance
            rows(i).RowHeight = m_intTrackCellHeight / (2 - (i - (6 + objBasicData.TOP_ROWS)) Mod 2)
        End If
    Next i
    Range(Cells(4 + objBasicData.TOP_ROWS, 1), Cells(objRace.NUMBER_ENROLLED * 2 + 19 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))).Interior.Color = objRace.TRACK_COLOUR 'Track colour
    With Range(Cells(4 + objBasicData.TOP_ROWS, 1), Cells(objRace.NUMBER_ENROLLED * 2 + 8 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))) 'Track font
        With .Font
            .name = "Arial"
            .size = m_intFontSize
            .Color = objOption.COL_TEXT
        End With
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    'In case of using starting gates
    If objRace.STARTING_GATE = "Y" Then
        'Draw the gates
            For i = 6 To (objRace.NUMBER_ENROLLED * 2 + 6) Step 2 'One gate for each starting place
                Range(Cells(i + objBasicData.TOP_ROWS, objBasicData.LEFT_COLS + 8), Cells(i + objBasicData.TOP_ROWS, objBasicData.LEFT_COLS + 13)).Interior.ColorIndex = 1
            Next i
            Range(Cells(6 + objBasicData.TOP_ROWS, objBasicData.LEFT_COLS + 13), Cells(objRace.NUMBER_ENROLLED * 2 + 6 + objBasicData.TOP_ROWS, objBasicData.LEFT_COLS + 13)).Interior.ColorIndex = 1 'Close the gates
        'Label the gates
            With Range(Cells(7 + objBasicData.TOP_ROWS, objBasicData.LEFT_COLS + 4), Cells(objRace.NUMBER_ENROLLED * 2 + 5 + objBasicData.TOP_ROWS, objBasicData.LEFT_COLS + 4))
                .Font.Color = objOption.COL_TEXT
                .Font.size = m_intFontSize
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
            End With
            For i = 1 To (objRace.NUMBER_ENROLLED) 'Gate numbers
                Cells(5 + 2 * i + objBasicData.TOP_ROWS, objBasicData.LEFT_COLS + 4).Value = GetText(g_arr_Text, "RACE010") & " " & i
            Next i
    End If
        
    'Display metres above and below the race track
    For i = objOption.METRES_DISPLAY To (objRace.METRES - 20) Step objOption.METRES_DISPLAY
        
        #If Debugging Then 'For debugging purposes: Vertical line at each marker position
            Range(Cells(5, i + objBasicData.LEFT_COLS + 11), Cells(45, i + objBasicData.LEFT_COLS + 11)).Interior.Color = objOption.COL_TEXT
        #End If
        
        With Cells(4 + objBasicData.TOP_ROWS, i + objBasicData.LEFT_COLS + 11) 'Above the track
            With .Font
                .name = "Arial"
                .Color = objOption.COL_TEXT
                .Bold = True
            End With
            .Value = i & GetText(g_arr_Text, "RACE008") '"m"
        End With
        With Cells(objRace.NUMBER_ENROLLED * 2 + 8 + objBasicData.TOP_ROWS, i + objBasicData.LEFT_COLS + 11) 'Below the track
            With .Font
                .name = "Arial"
                .Color = objOption.COL_TEXT
                .Bold = True
            End With
            .Value = i & GetText(g_arr_Text, "RACE008") '"m"
        End With
    Next i
    
    'Formatting: Horse names on the left
    With Range(Cells(6 + objBasicData.TOP_ROWS, 11), Cells(objRace.NUMBER_ENROLLED * 2 + 7 + objBasicData.TOP_ROWS, 12))
        .Font.Color = objRace.TRACK_COLOUR 'Track colour, so that the names are not visible yet
        .IndentLevel = 1 'Text indented
        .Font.size = m_intFontSize
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    'If chosen: Reveal individual race tactics
    If objOption.TACTICS And objOption.TACTICS_REVEAL_TAC Then
        With Columns(1)
            .HorizontalAlignment = xlLeft
            .Font.Color = objRace.TRACK_COLOUR ' "Hide" the text
            .Font.Bold = True
        End With
        Cells(objBasicData.TOP_ROWS + 5, 1).Value = GetText(g_arr_Text, "RACEOPT070") 'Caption
        For i = 1 To objRace.NUMBER_ENROLLED 'Race tactics of the horses
            If g_arr_varHorses(i, 0) = "START" Then
                With Cells(g_arr_varHorses(i, 3), 1)
                    .Font.name = "Courier New"
                    .NumberFormat = "@" 'Format as text
                    .Value = g_arr_varHorses(i, 25)
                End With
            End If
        Next i
        Columns(1).EntireColumn.AUTOFIT
    End If

    'If chosen: Formattings for the current speed according to the race tactics
    If objOption.TACTICS And objOption.TACTICS_REVEAL_CURR Then
        With Columns(10).Font
            .Color = objRace.TRACK_COLOUR
            .Bold = True
        End With
        Cells(objBasicData.TOP_ROWS + 5, 10).Value = GetText(g_arr_Text, "RACEOPT067") 'Caption
    End If
        
    'In case of a special race: Add the respective elements
    If objRace.TRACK_SURFACE = "M" Then Call DrawMudflats
    If objRace.SPECIAL = "PARTICULATES" Then Call DrawDust
        
    'Formatting: Finish area
    Columns(objRace.METRES + objBasicData.LEFT_COLS + 12).ColumnWidth = m_dblTrackCellWidth 'Width of the finish line
    Range(Cells(5 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 12), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 7 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 12)).Interior.ColorIndex = 56 'Colour of the finish line: dark grey
    With Cells(4 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 11) 'Race distance above the track
        With .Font
            .name = "Arial"
            .Color = objOption.COL_TEXT
            .Bold = True
        End With
        .Value = objRace.METRES & GetText(g_arr_Text, "RACE008")
    End With
    With Cells(objRace.NUMBER_ENROLLED * 2 + 8 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 11)  'Race distance below the track
        With .Font
            .name = "Arial"
            .Color = objOption.COL_TEXT
            .Bold = True
        End With
        .Value = objRace.METRES & GetText(g_arr_Text, "RACE008")
    End With
        
    'Formatting: Area behind the finish line
    Columns(objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)).ColumnWidth = 36
    Columns(objRace.METRES + objBasicData.LEFT_COLS + 15 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)).ColumnWidth = 9
    Range(Columns(objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), Columns(objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR) + 300)).ColumnWidth = m_dblTrackCellWidth 'Column width
    
    'Formatting: Horse names behind the finish line
    With Range(Cells(5 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), Cells(objRace.NUMBER_ENROLLED * 2 + 7 + (2 * objOption.SPEED_FACTOR) + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)))
        .IndentLevel = 1
        .Font.Color = objOption.COL_TEXT
        .Font.size = m_intFontSize
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    'Advertising area below the race track
    Range(rows(objRace.NUMBER_ENROLLED * 2 + 9 + objBasicData.TOP_ROWS), _
        rows(objRace.NUMBER_ENROLLED * 2 + 19 + objBasicData.TOP_ROWS)).EntireRow.RowHeight = m_intAdvertisingHeight 'Row height

    If objRace.ADVERTISING = "Y" Then
        Dim advPos As Integer 'Column position for the next ad
        
        'Clear the array with graphics data
        Erase m_arr_varTrackGraphics

        If replay Then
            m_arr_varTrackGraphics = g_arr_varReplay_RaceData(20, 1)
        Else
            Call GetTrackGraphicsData("ADVERTISEMENT") 'Get advertisements from the race sheet
        End If

        If Not Not m_arr_varTrackGraphics Then 'Only if the array contains at least one element
            advPos = objBasicData.LEFT_COLS + 12
            For i = 1 To UBound(m_arr_varTrackGraphics) 'Loop through the array which contains the ads
                If m_arr_varTrackGraphics(i) <> "" Then 'Skip if the element is empty
                    intPicColumn = GetColumn(g_wksPIC, m_arr_varTrackGraphics(i)) 'Find the column with the graphics data
                    If intPicColumn = 0 Then intPicColumn = GetColumn(g_wksPIC, "ADV_NOT_FOUND") 'Fallback
                    z = 3 'Set the pointer to the first colour code
                    For j = advPos To advPos + g_wksPIC.Cells(2, intPicColumn) - 1
                        If j >= objRace.METRES + objBasicData.LEFT_COLS + 12 Then Exit For 'Stop drawing behind the finish line
                        For k = objRace.NUMBER_ENROLLED * 2 + 9 + objBasicData.TOP_ROWS To objRace.NUMBER_ENROLLED * 2 + 19 + objBasicData.TOP_ROWS
                            Cells(k, j).Interior.Color = g_wksPIC.Cells(z, intPicColumn).Value

                            'Take the colour mode into account
                            Select Case g_strColourMode
                                Case "POPART"
                                    If Cells(k, j).Interior.Color > 0 And Cells(k, j).Interior.Color < 16777215 Then _
                                    Cells(k, j).Interior.Color = PopArtColour(Cells(k, j).Interior.Color)
                                Case "LSD"
                                    If Cells(k, j).Interior.Color > 0 And Cells(k, j).Interior.Color < 16777215 Then _
                                    Cells(k, j).Interior.Color = PopArtColour(Int((16777215 - 0 + 1) * Rnd + 0))
                                Case "SMARTIES"
                                    If Cells(k, j).Interior.Color > 0 And Cells(k, j).Interior.Color < 16777215 Then _
                                    Cells(k, j).Interior.Color = Int((16777215 - 0 + 1) * Rnd + 0)
                                Case "TV1960"
                                    Cells(k, j).Interior.Color = GreyToLong(CInt(RGBtoGrey(CLng(Cells(k, j).Interior.Color))))
                                Case "DARKMODE"
                                    Select Case Cells(k, j).Interior.Color
                                        Case 0 'Black: no change

                                        Case 16777215 'White: no change

                                        Case Else
                                            Cells(k, j).Interior.Color = DarkModeColour(Cells(k, j).Interior.Color)
                                    End Select
                                Case "24H"
                                    Cells(k, j).Interior.Color = DuskDawn(Cells(k, j).Interior.Color, Abs(22 * objOption.DAYLIGHT))
                            End Select
                            z = z + 1
                        Next k 'Next row
                    Next j 'Next column
                    advPos = advPos + g_wksPIC.Cells(2, intPicColumn) 'Column for the beginning of the next ad
                End If
            Next i
        End If
    End If
    
    'Spectator variables
    Dim intSpecSize(1 To 8) As Integer 'Size of the spectators
    Dim intSpecStart As Integer 'Position of the first spectator
    Dim intSpecFinish As Integer 'Position of the last spectator
    Dim lngSpecColor As Long 'Colour of the spectators

    intSpecStart = 12 + objBasicData.LEFT_COLS + 100
    
    intSpecSize(1) = 10 'Standing
    intSpecSize(2) = 11 'Standing
    intSpecSize(3) = 12 'Standing
    intSpecSize(4) = 8 'Sitting
    intSpecSize(5) = 10 'Sitting
    intSpecSize(6) = 11 'Sitting
    intSpecSize(7) = -4131 'Horizontal alignment: xlLeft
    intSpecSize(8) = -4108 'Vertical alignment: xlCenter
    
    Select Case objRace.RACE_ID
        Case "SPACE"
            lngSpecColor = vbGreen 'Alien green
            intNumberSpectators = objOption.SPECTATORS 'Spectators as chosen in the race options
        Case "CORONA2020", "CORONA2021Q1", "CORONA2021Q2", "CORONA2021Q3", "CORONA2021F"
            intNumberSpectators = 0 'No spectators allowed
            If objOption.SPECTATORS > 0 Then
                m_arr_strSpectators(0) = "X"
            Else
                m_arr_strSpectators(0) = ""
            End If
            m_arr_strSpectators(1) = GetText(g_arr_Text, "SPECT001")
        Case "DD20"
            intNumberSpectators = 4 'No common spectators allowed
            m_arr_strSpectators(0) = "X"
            m_arr_strSpectators(1) = GetText(g_arr_Text, "SPECT002")
        Case "GMP2018", "GMP2019FINAL"
            intNumberSpectators = 100 'Full house
            m_arr_strSpectators(0) = "X"
            m_arr_strSpectators(1) = GetText(g_arr_Text, "SPECT003")
        Case "DD17", "DD18", "DD19"
            intNumberSpectators = 100 'Full house
            m_arr_strSpectators(0) = "X"
            m_arr_strSpectators(1) = GetText(g_arr_Text, "SPECT004")
        Case "GMP2020Q6", "GMP2020SF1", "GMP2020SF2", "GMP2020FINAL"
            intNumberSpectators = 33 'Hamburg Coronavirus rules from July 1st 2020 https://www.hamburg.de/coronavirus/14031552/2020-06-30-sk-corona-aktuell/
            m_arr_strSpectators(0) = "X"
            m_arr_strSpectators(1) = GetText(g_arr_Text, "SPECT005")
        Case "GMP2020Q1", "GMP2020Q2", "GMP2020Q3", "GMP2020Q4", "GMP2020Q5"
            intNumberSpectators = 6 'Hamburg Coronavirus rules until June 30th 2020
            m_arr_strSpectators(0) = "X"
            m_arr_strSpectators(1) = GetText(g_arr_Text, "SPECT006")
        Case Else
            lngSpecColor = vbBlack
            intNumberSpectators = objOption.SPECTATORS 'Spectators as chosen in the race options
            m_arr_strSpectators(0) = "O"
    End Select
    Select Case g_strColourMode
        Case "DARKMODE"
            lngSpecColor = vbWhite
    End Select
        
    'Spectators standing above the race track
    If intNumberSpectators > 0 Then
        intSpecFinish = 12 + objBasicData.LEFT_COLS + objRace.METRES + objBasicData.AFTER_FIN_COLS
        
        'Prepare the speactator settings
        With Range(Cells(3 + objBasicData.TOP_ROWS, intSpecStart), Cells(3 + objBasicData.TOP_ROWS, intSpecFinish))

            .Font.Color = lngSpecColor 'Spectator colour
        End With

        'Populate with spectators
        For i = intSpecStart To intSpecFinish
            If Int(((100 / intNumberSpectators) - 1 + 1) * Rnd + 1) = 1 Then
                With Cells(3 + objBasicData.TOP_ROWS, i)
                    .Value = "i" 'Spectator (standing)
                    .Font.size = intSpecSize(Int((3 - 1 + 1) * Rnd + 1)) 'Child, adult...
                    .Font.Bold = Int((2 - 0 + 1) * Rnd + 0) 'Slim (0) or not (1 or 2)
                End With
            End If
        Next
    End If
    
    'Graphical elements like tibunes, food and beverage stands etc.
    'above the race track
    If objOption.TRIBUNES Then
    
        'Clear the array with graphics data
        Erase m_arr_varTrackGraphics
        
        'Element variables
        Dim rep As Integer 'Number of repetitions of an element
        Dim graphicsPos As Integer 'Column position for the next element
        Dim lngTreeColour(1 To 5) As Long 'Tree colours
        Dim lngMountainColour(1 To 4) As Long 'Mountain colours
        Dim lngSoilColour(1 To 3) As Long 'Soil colours
        Dim intTextSize(1 To 3) As Integer
        Dim lngSpecialCounter As Long 'Counter for secial purposes

        intTextSize(1) = 11
        intTextSize(2) = 10
        intTextSize(3) = 18

        graphicsPos = objBasicData.LEFT_COLS + 12 'Set the pointer to the starting line
        
        'Define colours for random calculations
        lngTreeColour(1) = 2315831 'Dark green (leaves)
        lngTreeColour(2) = 3506772 'Light green (leaves)
        lngTreeColour(3) = 3689263 'Medium green (conifer)
        lngTreeColour(4) = 2503712 'Dark green (conifer)
        lngTreeColour(5) = 4743485 'Light green (conifer)
        lngMountainColour(1) = 15196117 'Light grey (rock)
        lngMountainColour(2) = 14277081 'Very light grey (rock)
        lngMountainColour(3) = 15652797 'Light blue (glacier)
        lngMountainColour(4) = 15983569  'Very light blue (glacier)
        lngSoilColour(1) = 3368601 'Medium brown
        lngSoilColour(2) = 4231094 'Light brown
        lngSoilColour(3) = 2378094 'Dark brown

        If replay Then
            m_arr_varTrackGraphics = g_arr_varReplay_RaceData(22, 1)
        Else
            Call GetTrackGraphicsData("TRACK GRAPHICS") 'Get element sequence from the race sheet
        End If

        If Not Not m_arr_varTrackGraphics Then 'Only if the array contains at least one element
            k = 0
            For i = 1 To UBound(m_arr_varTrackGraphics) 'Loop through the array which contains the elements
                If m_arr_varTrackGraphics(i) <> "" Then 'Skip if the element is empty
                    intPicColumn = GetColumn(g_wksPIC, m_arr_varTrackGraphics(i)) 'Find the column with the graphics data
                    If intPicColumn = 0 Then intPicColumn = GetColumn(g_wksPIC, "TRACKGRAPHICS_NOT_FOUND") 'Fallback
                    If k + g_wksPIC.Cells(2, intPicColumn) >= 12 + objBasicData.LEFT_COLS + objRace.METRES + objBasicData.AFTER_FIN_COLS Then Exit For 'Stop drawing behind the finish line
                    
                    'Consider repetitions
                    For rep = 1 To g_wksPIC.Cells(4, intPicColumn)
                    
                        z = 6 'Set the pointer to the first colour code
                        lngSpecialCounter = 0 'Reset the counter for special purposes
                    
                        For j = objBasicData.TOP_ROWS + 1 To objBasicData.TOP_ROWS + 3
                            For k = graphicsPos To graphicsPos + g_wksPIC.Cells(3, intPicColumn) - 1
                                If Not IsEmpty(g_wksPIC.Cells(z, intPicColumn)) Then 'Draw only if a value is found
                                    Cells(j, k).Interior.Color = g_wksPIC.Cells(z, intPicColumn).Value
                                End If
                                'Draw graphical elements
                                Select Case g_wksPIC.Cells(5, intPicColumn)
                                    Case "TRIBUNE"
                                        'Prepare tribunes
                                        With Cells(j, k)
                                            .ClearContents 'No spectators standing in front of the tribune
                                            .Font.Italic = True
                                            .HorizontalAlignment = intSpecSize(7)
                                            .VerticalAlignment = intSpecSize(8)
                                            .Font.Color = lngSpecColor
                                        End With
                                        'Populate tribunes with sitting spectators
                                        If intNumberSpectators > 0 And Cells(j, k).Interior.ColorIndex = 15 Then 'Seat found
                                            If Int(((100 / intNumberSpectators) - 1 + 1) * Rnd + 1) = 1 Then
                                                With Cells(j, k)
                                                    .Value = "i" 'Spectator (sitting)
                                                    .Font.size = intSpecSize(Int((6 - 4 + 1) * Rnd + 4)) 'Child, adult...
                                                    .Font.Bold = Int((2 - 0 + 1) * Rnd + 0) 'Slim (0) or not (1 or 2)
                                                End With
                                            End If
                                        End If
                                    Case "TREE", "MOUNTAINS"
                                        'Trees
                                        If Cells(j, k).Interior.Color = lngTreeColour(1) Then  'Dark green leave found
                                            If (Int((10 - 1 + 1) * Rnd + 1)) = 10 Then Cells(j, k).Interior.Color = lngTreeColour(2) '1 out of 10 leaves light green
                                        End If
                                        If Cells(j, k).Interior.Color = lngTreeColour(3) Then  'Medium green conifer found
                                            If (Int((2 - 1 + 1) * Rnd + 1)) = 2 Then Cells(j, k).Interior.Color = lngTreeColour(4) '1 out of 2 conifers dark green
                                            If (Int((6 - 1 + 1) * Rnd + 1)) = 6 Then Cells(j, k).Interior.Color = lngTreeColour(5) '1 out of 6 conifers light green
                                        End If
                                        'Mountains
                                        If Cells(j, k).Interior.Color = lngMountainColour(1) Then  'Mountain (rock) found
                                            If (Int((6 - 1 + 1) * Rnd + 1)) = 6 Then Cells(j, k).Interior.Color = lngMountainColour(2) '1 out of 6 lighter grey
                                        End If
                                        If Cells(j, k).Interior.Color = lngMountainColour(3) Then  'Mountain (glacier) found
                                            If (Int((6 - 1 + 1) * Rnd + 1)) = 6 Then Cells(j, k).Interior.Color = lngMountainColour(4) '1 out of 6 lighter blue
                                        End If
                                    Case "MOUND"
                                        'Soil
                                        If Cells(j, k).Interior.Color = lngSoilColour(1) Then 'Soil found
                                            If (Int((6 - 1 + 1) * Rnd + 1)) = 6 Then Cells(j, k).Interior.Color = lngSoilColour(2) '1 out of 6 light brown
                                        End If
                                        If Cells(j, k).Interior.Color = lngSoilColour(1) Then  'Soil found
                                            If (Int((3 - 1 + 1) * Rnd + 1)) = 3 Then Cells(j, k).Interior.Color = lngSoilColour(3) '1 out of 3 dark brown
                                        End If
                                    Case "HOTEL"
                                        Cells(j, k).ClearContents 'No spectators standing in front of the hotel
                                        If j = objBasicData.TOP_ROWS + 1 And k = graphicsPos + 2 Then
                                            With Cells(j, k)
                                                .VerticalAlignment = xlCenter
                                                .Value = GetTrackGraphicsText(g_wksPIC.Cells(5, intPicColumn))
                                                With .Font
                                                    .Color = GetTrackGraphicsFontCol(g_wksPIC.Cells(5, intPicColumn))
                                                    .Bold = True
                                                    .size = intTextSize(1)
                                                End With
                                            End With
                                        End If
                                    Case "ENTRANCE", "CHIPS", "HOTDOGS", "MATJES", "FISHANDCHIPS", _
                                            "BEER", "BETS", "ROESTI", "FONDUE", "LIMO", "COLA", "CURRYWURST"
                                        If j = objBasicData.TOP_ROWS + 2 And k = graphicsPos + CInt((g_wksPIC.Cells(2, intPicColumn) / 2)) - 1 Then
                                            With Cells(j, k)
                                                .HorizontalAlignment = xlCenter
                                                .VerticalAlignment = xlCenter
                                                .Value = GetTrackGraphicsText(g_wksPIC.Cells(5, intPicColumn))
                                                With .Font
                                                    .Color = GetTrackGraphicsFontCol(g_wksPIC.Cells(5, intPicColumn))
                                                    .Bold = True
                                                    .size = intTextSize(2)
                                                End With
                                            End With
                                        End If
                                    Case "BUILDING_SAP"
                                        If Cells(j, k).Interior.Color = 15445507 Then
                                            lngSpecialCounter = lngSpecialCounter + 1
                                            If lngSpecialCounter = 5 Then
                                                With Cells(j, k)
                                                    .HorizontalAlignment = xlCenter
                                                    .VerticalAlignment = xlCenter
                                                    .Value = "SAP"
                                                    With .Font
                                                        .name = "Arial Black"
                                                        .Color = vbWhite
                                                        .Bold = True
                                                        .size = intTextSize(3)
                                                    End With
                                                End With
                                            End If
                                        End If
                                    Case "BUILDING_OTTO"
                                        If Cells(j, k).Interior.Color = 9349808 Then
                                            lngSpecialCounter = lngSpecialCounter + 1
                                            If lngSpecialCounter = 7 Then
                                                With Cells(j, k)
                                                    .HorizontalAlignment = xlCenter
                                                    .VerticalAlignment = xlCenter
                                                    .Value = "OTTO"
                                                    With .Font
                                                        .name = "Franklin Gothic Heavy"
                                                        .Color = 1901268
                                                        .size = intTextSize(1)
                                                    End With
                                                End With
                                            End If
                                        End If
                                End Select
    
                                'Take the colour mode into account
                                Select Case g_strColourMode
                                    Case "POPART"
                                        If Cells(j, k).Interior.Color > 0 And Cells(j, k).Interior.Color < 16777215 Then _
                                        Cells(j, k).Interior.Color = PopArtColour(Cells(j, k).Interior.Color)
                                    Case "LSD"
                                        If Cells(j, k).Interior.Color > 0 And Cells(j, k).Interior.Color < 16777215 Then _
                                        Cells(j, k).Interior.Color = PopArtColour(Int((16777215 - 0 + 1) * Rnd + 0))
                                    Case "SMARTIES"
                                        If Cells(j, k).Interior.Color > 0 And Cells(j, k).Interior.Color < 16777215 Then _
                                        Cells(j, k).Interior.Color = Int((16777215 - 0 + 1) * Rnd + 0)
                                    Case "TV1960"
                                        Cells(j, k).Interior.Color = GreyToLong(CInt(RGBtoGrey(CLng(Cells(j, k).Interior.Color))))
                                    Case "DARKMODE"
                                        Select Case Cells(j, k).Interior.Color
                                            Case 0  'Black: no change
    
                                            Case 16777215  'White: turn black
                                                Cells(j, k).Interior.Color = 0
                                            Case Else
                                                Cells(j, k).Interior.Color = DarkModeColour(Cells(j, k).Interior.Color)
                                        End Select
                                    Case "24H"
                                        Select Case Cells(j, k).Interior.Color
                                            Case objOption.DAYLIGHT_COL 'Heaven: no change
                                            
                                            Case 16777215  'White: turn heaven
                                                Cells(j, k).Interior.Color = objOption.DAYLIGHT_COL
                                            Case Else
                                                Cells(j, k).Interior.Color = DuskDawn(Cells(j, k).Interior.Color, Abs(22 * objOption.DAYLIGHT))
                                        End Select
                                End Select
                                z = z + 1 'Next colour code
                            Next k 'Next column
                        Next j 'Next row
                        
                        graphicsPos = graphicsPos + g_wksPIC.Cells(3, intPicColumn) 'Column for the beginning of the next ad
                    Next rep
                End If
            Next i
        End If
    End If

    If objOption.AUTOFIT Then Call AutoZoom("Race")
    
    'De-activate screen updating
    Application.ScreenUpdating = False
    
    'Place the cursor far away
    Cells(100 + objBasicData.TOP_ROWS, 1).Select
    
    'Activate screen updating
    Application.ScreenUpdating = True
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "DrawRaceTrack()")
    Call basAuxiliary.CodeCrash
End Sub

'Get the typical description for the country in which the race takes place
'(e.g. "Chips" in England, "French fries" in the USA)
Private Function GetTrackGraphicsText(element As String) As String

    Dim raceLanguage As String 'Language spoken in this country
    
    Select Case objRace.COUNTRY_CODE
        Case "DEU", "CHE"
            raceLanguage = "DE"
        Case "USA", "AUS"
            raceLanguage = "US"
        Case Else
            raceLanguage = "EN"
    End Select
    
    GetTrackGraphicsText = g_wksTEXT.Cells(GetRow(g_wksTEXT, element), GetColumn(g_wksTEXT, raceLanguage)).Value

End Function

'Get the font colour for the element
Private Function GetTrackGraphicsFontCol(element As String) As Long
    Select Case element
        Case "ENTRANCE", "CHIPS", "MATJES", "BEER", "BETS", "FISHANDCHIPS", _
                "LIMO", "COLA"
            GetTrackGraphicsFontCol = vbWhite
        Case "HOTDOGS", "HOTEL", "FONDUE"
            GetTrackGraphicsFontCol = vbBlack
        Case "ROESTI"
            GetTrackGraphicsFontCol = vbYellow
        Case "CURRYWURST"
            GetTrackGraphicsFontCol = vbRed
    End Select
End Function

'Track extensions for a mudflats race
Private Sub DrawMudflats()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Call basAuxiliary.ActivateRaceSheet
    
    'For debugging purposes: Lugworm population
    #If Debugging Then
        Dim lngLug1 As Long, lngLug2 As Long
    #End If
        
    'Variables for puddles
    Dim intPuddleFrequency As Integer
    Dim intPuddleLength As Integer
    Dim intPuddleWidth As Integer

    'Variables for lugworms
    Dim intLugwormFrequency As Integer
    Dim intLugwormShape As Integer
    
    Dim c As Integer 'Column
        
    'Draw lugworms
    If objOption.LUGWORMS > 0 Then
        objOption.LUGWORM_COL = 2770764 'Colour of the lugworms in the Wadden Sea
        c = basAuxiliary.GetColumn(g_wksPIC, "LUGWORMS") 'Get the column with the lugworm characters
        ReDim m_arr_lngLugworms(1 To (g_wksPIC.Cells(rows.count, c).End(xlUp).row - 1))
        
        'Read the character codes for the lugworm characters from the worksheet "PIC"
        For i = 1 To UBound(m_arr_lngLugworms)
            m_arr_lngLugworms(i) = g_wksPIC.Cells(i + 1, c)
        Next i

        For i = (5 + objBasicData.TOP_ROWS) To (objRace.NUMBER_ENROLLED * 2 + 7 + objBasicData.TOP_ROWS) 'Loop through the rows
            For j = objBasicData.LEFT_COLS + 15 To objRace.METRES + objBasicData.LEFT_COLS + 7 'Loop through the columns
                intLugwormFrequency = Int(((100 / objOption.LUGWORMS) - 1 + 1) * Rnd + 1) 'Lugworm or no lugworm
                If intLugwormFrequency = 1 Then
                    intLugwormShape = Int((UBound(m_arr_lngLugworms) - 1 + 1) * Rnd + 1) 'Shape of the lugworm
                    
                    With Cells(i, j) 'Draw the lugworm
                        .Font.Color = objOption.LUGWORM_COL 'Colour
                        .Value = ChrW(m_arr_lngLugworms(intLugwormShape)) 'Shape
                    End With
                    
                    #If Debugging Then
                        lngLug1 = lngLug1 + 1 'Count the number of lugworms
                    #End If
                End If
                #If Debugging Then
                    lngLug2 = lngLug2 + 1 'Count the total number of cells
                #End If
            Next j
        Next i
        
        #If Debugging Then 'Lugworm density (%)
            Debug.Print vbNewLine & lngLug1 _
                & " lugworms (population " & Round(lngLug1 / lngLug2 * 100, 1) & "%)"
        #End If
    End If

    'Draw Puddles
    If objOption.TIDE > 0 Then
        objOption.PUDDLE_COL = 10791854 'Colour of the puddles in the Wadden Sea
        
        For i = (5 + objBasicData.TOP_ROWS) To (objRace.NUMBER_ENROLLED * 2 + 7 + objBasicData.TOP_ROWS) 'Loop through the rows
            For j = objBasicData.LEFT_COLS + 15 To objRace.METRES + objBasicData.LEFT_COLS + 7 'Loop through the columns
                intPuddleFrequency = Int(((100 / objOption.TIDE) - 1 + 1) * Rnd + 1) 'Puddle frequency
                
                If intPuddleFrequency = 1 Then
                    intPuddleLength = Int((10 - 1 + 1) * Rnd + 1) 'Puddle length (number of columns)
                    intPuddleWidth = Int((2 - 1 + 1) * Rnd + 1) 'Puddle width (number of rows)
                    
                    'Draw the puddle
                    With Range(Cells(i, j), Cells(i + intPuddleWidth - 1, j + intPuddleLength - 1))
                        .Interior.Color = objOption.PUDDLE_COL
                        .Font.Color = objOption.PUDDLE_COL
                        .Value = "|" 'Cell content that marks the cell as a puddle (for technical purposes). Not visible as the font colour matches the cell colour
                    End With
                    
                    #If Debugging Then
                        Range(Cells(i, j), Cells(i + intPuddleWidth - 1, j + intPuddleLength - 1)) _
                            .Font.Color = vbBlack 'Make the vertical bar characters visible
                    #End If

                End If

            Next j
        Next i
    End If

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "DrawMudflats()")
    Call basAuxiliary.CodeCrash
End Sub

'Draw a layer of particulates
Private Sub DrawDust()

    Dim intArr() As Variant
    
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    intArr() = Array(18, 17, 4, 16, 9, 10) 'Pattern values
    objOption.PARTICULATES_PATTERN = intArr(objOption.PARTICULATES_SLIDER)
    
    Call basAuxiliary.ActivateRaceSheet

    Cells.Interior.Pattern = objOption.PARTICULATES_PATTERN 'Lay a pattern on all cells
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "DrawDust()")
    Call basAuxiliary.CodeCrash
End Sub

'Coronavirus testing
Private Sub CoronavirusTest()
    If objRace.SPECIAL = "CORONAVIRUSTEST" Then frmCoronavirusTest.show
End Sub

'Display horse names during the race
Private Sub DrawHorseNames()
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Call basAuxiliary.ActivateRaceSheet
    
    'Display the horse names and starting numbers on the left if selected in the race options
    If objOption.NAMES_LEFT Then
        For i = 1 To objRace.NUMBER_ENROLLED 'Loop through the horses
            If g_arr_varHorses(i, 0) = "START" Then 'Draw only if the horse is starting
                Cells(g_arr_varHorses(i, 3), 11).Value = "#" & g_arr_varHorses(i, 11) 'Number
                Cells(g_arr_varHorses(i, 3), 12).Value = g_arr_varHorses(i, 1) 'Name
                'As the font colour matches the track colour the horse numbers and names will not be visible yet
                #If Debugging Then 'For debugging purposes: Change the font colour so that the text is visible
                    Range(Cells(g_arr_varHorses(i, 3), 11), Cells(g_arr_varHorses(i, 3), 12)).Font.ColorIndex = 16 'Grey
                #End If
            End If
        Next i
    End If
    
    'Display the horse names and starting numbers behind the finish line if selected in the race options
    If objOption.NAMES_FINISH Then
        For j = 1 To objRace.NUMBER_ENROLLED
            If g_arr_varHorses(j, 0) = "START" Then 'Only if the horse is starting
                Cells(g_arr_varHorses(j, 3), objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)).Value _
                    = g_arr_varHorses(j, 1)                   'Horse name
            End If
        Next j
    End If
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "DrawHorseNames()")
    Call basAuxiliary.CodeCrash
End Sub

'Prepare the race sheet for displaying the current momentum speed bars of each horse
Private Sub MomentumFormattings_Bars()
    
    #If Debugging Then
        'Expand the column width to see the last speed values
        Range(Columns(2), Columns(9)).ColumnWidth = 4
    #End If
    
    Columns(13).ColumnWidth = 12
    With Cells(5 + objBasicData.TOP_ROWS, 13)
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Font.Color = objRace.TRACK_COLOUR 'Hide the text
        .Value = GetText(g_arr_Text, "RACEOPT078")
    End With

    'Speed bar
    With Range(Cells(6 + objBasicData.TOP_ROWS, 13), Cells(objRace.NUMBER_ENROLLED * 2 + 6 + objBasicData.TOP_ROWS, 13))
        .FormatConditions.AddDatabar
        .FormatConditions(.FormatConditions.count).ShowValue = False 'Show no values, just the bars
        With .FormatConditions(1)
            .BarColor.Color = objRace.TRACK_COLOUR
            .BarBorder.Type = xlDataBarBorderSolid 'xlDataBarBorderNone
            .BarFillType = xlDataBarFillGradient 'xlDataBarFillSolid
        End With
    End With
End Sub

'Prepare the race sheet for displaying the current momentum speed icons of each horse
Private Sub MomentumFormattings_Icons()
    
    Columns(13).ColumnWidth = 12
    With Cells(5 + objBasicData.TOP_ROWS, 13)
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Font.Color = objRace.TRACK_COLOUR 'Hide the text
        .Value = GetText(g_arr_Text, "RACEOPT078")
    End With

    With Range(Cells(6 + objBasicData.TOP_ROWS, 13), Cells(objRace.NUMBER_ENROLLED * 2 + 6 + objBasicData.TOP_ROWS, 13))
        .FormatConditions.AddIconSetCondition
        .FormatConditions(.FormatConditions.count).SetFirstPriority
        With .FormatConditions(1)
            .ReverseOrder = False
            .ShowIconOnly = True  'Show no values, just the icons
            .IconSet = ActiveWorkbook.IconSets(xl3Triangles)
        End With
        With .FormatConditions(1).IconCriteria(2)
            .Type = xlConditionValuePercent
            .Value = 10 'Top 10%
            .Operator = 7
        End With
        With .FormatConditions(1).IconCriteria(3)
            .Type = xlConditionValuePercent
            .Value = 90 'Last 10%
            .Operator = 7
        End With
    End With
End Sub

'Race information pop-up
Private Sub RaceInfoPopup()
    With frmRaceInfo
        'Place the pop-up in the upper left corner
        .StartUpPosition = 0
        .top = ActiveWindow.top + 40
        .left = ActiveWindow.left + 40
        .Height = 72
        .width = 225
        .BackColor = objOption.RACE_INFO_COL_B
        .caption = GetText(g_arr_Text, "USERFORM006")

        'Label with the race name and year
        With .lbl_RI1
            .BackColor = objOption.RACE_INFO_COL_B
            .ForeColor = objOption.RACE_INFO_COL_F
            .caption = objRace.RACE_NAME & " " & objRace.RACE_YEAR
            .Font.size = 8.5
            .Font.Bold = True
            .AutoSize = True
        End With
        
        'Label with the race distance
        With .lbl_RI2
            .BackColor = objOption.RACE_INFO_COL_B
            .ForeColor = objOption.RACE_INFO_COL_F
            .Font.size = 12
            .caption = GetText(g_arr_Text, "RACE024") & ": " & objRace.METRES & GetText(g_arr_Text, "RACE008")
            .AutoSize = True
        End With
        
        'Add a progress bar at runtime if selected in the race options
        If objOption.RACE_INFO_PROGRESS Then
        
            'Label which serves as a frame for the progress bar
            Set g_objLabel = .Controls.Add("Forms.Label.1", , True)
            With g_objLabel
                .name = "lbl_RI3a_dyn" 'ID of the new object
                .Font.name = "Tahoma"
                .Font.size = 8
                .left = 6
                .top = frmRaceInfo.Height - 25
                .width = 200
                .Height = 12
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = objOption.RACE_INFO_COL_F
                .ForeColor = objOption.RACE_INFO_COL_F
                .BackColor = objOption.RACE_INFO_COL_B
                'Display the race distance permanently at the right edge
                .TextAlign = fmTextAlignRight
                .caption = objRace.METRES
            End With
            
            'Label for the progress bar itself
            Set g_objLabel = .Controls.Add("Forms.Label.1", , True)
            With g_objLabel
                .name = "lbl_RI3b_dyn" 'ID of the new object
                .Font.name = "Tahoma"
                .Font.size = 8
                .left = 6
                .top = frmRaceInfo.Height - 25
                .width = 0
                .Height = 12
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = objOption.RACE_INFO_COL_F
                .ForeColor = objOption.RACE_INFO_COL_B
                .BackColor = objOption.RACE_INFO_COL_F
                .TextAlign = fmTextAlignLeft
            End With
            
            'Adjust the height of the info pop-up
            frmRaceInfo.Height = frmRaceInfo.Height + g_objLabel.Height + 6
        End If

        'Add a label for the leading horse at runtime if selected in the race options
        If objOption.RACE_INFO_LEADER Then
        
            'Label for the text "The current leader is..."
            Set g_objLabel = .Controls.Add("Forms.Label.1", , True)
            With g_objLabel
                .name = "lbl_RI4a_dyn" 'ID of the new object
                .Font.size = 12
                .left = 6
                .top = frmRaceInfo.Height - 25
                .width = 200
                .Height = 18
                .ForeColor = objOption.RACE_INFO_COL_F
                .TextAlign = fmTextAlignLeft
                .caption = ""
            End With
            
            'Adjust the height of the info pop-up
            frmRaceInfo.Height = frmRaceInfo.Height + g_objLabel.Height
            
            'Label for the name of the leader
            Set g_objLabel = .Controls.Add("Forms.Label.1", , True)
            With g_objLabel
                .name = "lbl_RI4b_dyn"
                .Font.size = 12
                .left = 6
                .top = frmRaceInfo.Height - 25
                .width = 200
                .Height = 18
                .ForeColor = objOption.RACE_INFO_COL_F
                .TextAlign = fmTextAlignCenter
                .caption = ""
            End With
            
            'Adjust the height of the info pop-up once more
            frmRaceInfo.Height = frmRaceInfo.Height + g_objLabel.Height + 6
        End If

        .show (vbModeless) 'Show the pop-up
    End With
End Sub

'Welcome message when a race begins
Private Sub RaceWelcome()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
'    'A MessageBox object cannot handle unicode, so for example Cyrillic characters are displayed as question marks
'    'Comment in the following lines and switch to Bulgarian language to experience the effect...
'    MsgBox GetText(g_arr_Text, "RACE001") & " " & GetText(g_arr_Text, "RACE006") & " " & objRace.TRACK_LOCATION & " (" & objRace.COUNTRY_NAME & "). " & vbNewLine & vbNewLine _
'            & GetText(g_arr_Text, "RACE003") & ": " & objRace.RACE_NAME & " " & GetText(g_arr_Text, "RACE007") & " " & objRace.METRES & " " & GetText(g_arr_Text, "RACE009") & "." & vbNewLine _
'            & GetText(g_arr_Text, "RACE004") & " " & objRace.NUMBER_STARTING & " " & g_arr_Grammar(4) & ".", , g_c_tool
    
    'Compile a message text that can be used in a pop-up as well as for voice output
    Dim messagetext As String
    messagetext = GetText(g_arr_Text, "RACE001") & " " & GetText(g_arr_Text, "RACE006") & " " & objRace.TRACK_LOCATION & " (" & objRace.COUNTRY_NAME & "). " & vbNewLine & vbNewLine _
            & GetText(g_arr_Text, "RACE003") & ": " & objRace.RACE_NAME & " " & GetText(g_arr_Text, "RACE007") & " " & objRace.METRES & " " & GetText(g_arr_Text, "RACE009") & "." & vbNewLine _
            & GetText(g_arr_Text, "RACE004") & " " & objRace.NUMBER_STARTING & " " & g_arr_Grammar(4) & "."
    If objOption.SPEECH Then Call SpeechOut(messagetext) 'Voice output if selected
    Call ShowMessagePopup(g_c_tool, messagetext, enumButton.OK, vbModal) 'Show a pop-up
    
    If m_arr_strSpectators(0) = "X" Then
        If objOption.SPEECH Then Call SpeechOut(m_arr_strSpectators(1)) 'Voice output if selected
        Call ShowMessagePopup(g_c_tool, m_arr_strSpectators(1), enumButton.OK, vbModal) 'Show a pop-up
    End If
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RaceWelcome()")
    Call basAuxiliary.CodeCrash
End Sub

'Put the horses in the starting gates
Private Sub StartingGrid()
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Call basAuxiliary.ActivateRaceSheet
    
    If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:02")) 'Delay
    
    For i = 1 To objRace.NUMBER_ENROLLED 'Loop through the horses
        If g_arr_varHorses(i, 0) = "START" Then 'Paint only horses that will run
            Call PaintHorse(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7, g_arr_varHorses(i, 2))
        End If
    Next i
    
    'Only in case of a tactical race
    If objOption.TACTICS And objOption.TACTICS_REVEAL_TAC Then
        Columns(1).Font.Color = objOption.COL_TEXT ' "Reveal" the text
    End If
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "StartingGrid()")
    Call basAuxiliary.CodeCrash
End Sub

'Presentation of the horses with numbers and names
Private Sub RacePresentation(test As Boolean, replay As Boolean)
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Call basAuxiliary.ActivateRaceSheet
    
    If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:02")) 'Delay

    If Not replay Then
        Application.DisplayCommentIndicator = xlCommentAndIndicator 'Display comments and indicators at all times
        
        'Display a comment for each horse
        For i = 1 To objRace.NUMBER_ENROLLED
            If g_arr_varHorses(i, 0) = "START" Then
                With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4))
                    'Extend the comment field if the horse is favourite or in focus
                    If i = m_byteFavourite(1) And objRace.RACE_ID <> "SPACE" Then
                        If objOption.FOCUSED_RUN = enumCamera.focus_horse And g_arr_varHorses(i, 11) = objOption.FOCUSED_NR Then 'Favourite and in focus
                            .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) _
                                & " (" & GetText(g_arr_Text, "RACE011") & ") >> " & GetText(g_arr_Text, "RACE012") 'Horse number, name and "(favourite) >> in focus"
                        Else 'Favourite but not in focus
                            .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) _
                                & " (" & GetText(g_arr_Text, "RACE011") & ")" 'Horse number, name and "(favourite)"
                        End If
                    ElseIf objOption.FOCUSED_RUN = enumCamera.focus_horse And g_arr_varHorses(i, 11) = objOption.FOCUSED_NR Then 'In focus but no favourite
                        .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) _
                            & " >> " & GetText(g_arr_Text, "RACE012") 'Horse number, name and ">> in focus"
                    Else 'No favourite and not in focus
                        .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) 'Horse number and name
                    End If
                    .Comment.Shape.TextFrame.Characters.Font.size = m_intFontSize 'Font size
                    .Comment.Shape.TextFrame.AutoSize = True 'Resize the comment field for perfect fit
                End With
                
                'In case of a Focused Run: Highight and draw a yellow dashed frame around the horse on focus
                If objOption.FOCUSED_RUN = enumCamera.focus_horse Then
                    If g_arr_varHorses(i, 11) = objOption.FOCUSED_NR Then
                        Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
                            Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)) _
                            .BorderAround ColorIndex:=44, LineStyle:=xlDash, Weight:=xlThick
                        Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)) _
                            .Comment.Shape.Fill.ForeColor.RGB = RGB(255, 204, 0) 'Yellow background
                    End If
                End If
                
                'Highlight the favourite horse
                If i = m_byteFavourite(1) And objRace.RACE_ID <> "SPACE" Then
                    Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)) _
                        .Comment.Shape.Fill.ForeColor.RGB = RGB(255, 0, 0) 'Red background
                End If
            End If
        Next i
    
        'Announce the three favourite horses
        If Not test And objOption.ANNOUNCE_FAV And objRace.RACE_ID <> "SPACE" Then
            Dim messagetext As String
            messagetext = GetText(g_arr_Text, "RACE013") & " " & g_arr_varHorses(m_byteFavourite(1), 1) & _
                        " (#" & g_arr_varHorses(m_byteFavourite(1), 11) & ")." & vbNewLine & vbNewLine & _
                        GetText(g_arr_Text, "RACE015") & " " & g_arr_varHorses(m_byteFavourite(2), 1) & " (#" _
                        & g_arr_varHorses(m_byteFavourite(2), 11) & ") " & _
                        GetText(g_arr_Text, "RACE017") & " " & g_arr_varHorses(m_byteFavourite(3), 1) & " (#" & _
                        g_arr_varHorses(m_byteFavourite(3), 11) & ") " & GetText(g_arr_Text, "RACE018") & "."
            
            If objOption.SPEECH Then Call SpeechOut(messagetext) 'Voice output if selected
            Call ShowMessagePopup(objRace.RACE_NAME & " " & objRace.RACE_YEAR, _
                messagetext, enumButton.OK, vbModal)
        End If
        
        'In case of a Focused Run: announce the focused horse
        If Not test And objOption.FOCUSED_RUN = enumCamera.focus_horse Then
            For i = 1 To UBound(g_arr_varHorses)
                If g_arr_varHorses(i, 11) = objOption.FOCUSED_NR Then
                    messagetext = GetText(g_arr_Text, "RACE021") & " " & g_arr_varHorses(i, 1) & " (#" & g_arr_varHorses(i, 11) & ")."
                    If objOption.SPEECH Then Call SpeechOut(messagetext) 'Voice output if selected
                    Call ShowMessagePopup(GetText(g_arr_Text, "RACEOPT026"), messagetext, _
                        enumButton.OK, vbModal)
                    Exit For
                End If
            Next i
        End If
                    
        'Turn off all cell comments (to hide the horse names)
        Application.DisplayCommentIndicator = xlNoIndicator
        
    End If
        
    'Show the horse colours on the left edge if selected in the race options
    If objOption.COLOURS_LEFT Then
        For i = 1 To objRace.NUMBER_ENROLLED
            If g_arr_varHorses(i, 0) = "START" Then
                Call PaintHorse(g_arr_varHorses(i, 3), 2, g_arr_varHorses(i, 2))
            End If
        Next i
    End If

    'Show the horse names and numbers at the start if selected in the race options
    Range(Columns(11), Columns(12)).Font.Color = objOption.COL_TEXT 'Change the font colour so that the names are visible
    
    'Mark the favourite on the left if selected in the race options
    If objOption.HIGHLIGHT_FAV And objRace.RACE_ID <> "SPACE" And Not replay Then
        Range(Cells(g_arr_varHorses(m_byteFavourite(1), 3), 11), Cells(g_arr_varHorses(m_byteFavourite(1), 3), 12)) _
            .Interior.Color = 255 'Red background
        'Show the number and name of the horse on the left (in case it does not already exist)
        Cells(g_arr_varHorses(m_byteFavourite(1), 3), 11).Value = "#" & g_arr_varHorses(m_byteFavourite(1), 11) 'Horse number
        Cells(g_arr_varHorses(m_byteFavourite(1), 3), 12).Value = g_arr_varHorses(m_byteFavourite(1), 1) _
                & " (" & GetText(g_arr_Text, "RACE011") & ")" 'Horse name
    End If
        
    'In case of a Focused Run: adapt the frame around the focused horse
    If objOption.FOCUSED_RUN = enumCamera.focus_horse Then
        For i = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(i, 11) = objOption.FOCUSED_NR Then
                'Delete the frame around the horse
                Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
                    Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)) _
                    .Borders.LineStyle = xlLineStyleNone
                'Highlight the focused horse on the left if selected in the race options
                If objOption.HIGHLIGHT_FOC Then
                    'Draw a yellow dashed frame around the horse name and number
                    Range(Cells(g_arr_varHorses(i, 3), 11), Cells(g_arr_varHorses(i, 3), 12)) _
                        .BorderAround ColorIndex:=44, LineStyle:=xlDash, Weight:=xlThick
                    'Show the number and name of the horse on the left (in case it does not already exist)
                    Cells(g_arr_varHorses(i, 3), 11).Value = "#" & g_arr_varHorses(i, 11) 'Horse number
                    If IsEmpty(Cells(g_arr_varHorses(i, 3), 12).Value) Then
                        Cells(g_arr_varHorses(i, 3), 12).Value = g_arr_varHorses(i, 1) _
                                & " >> " & GetText(g_arr_Text, "RACE012") ' ">> in focus"
                    Else
                        Cells(g_arr_varHorses(i, 3), 12).Value = Cells(g_arr_varHorses(i, 3), 12).Value _
                                & " >> " & GetText(g_arr_Text, "RACE012") ' "(favourite) >> in focus"
                    End If

                End If
                Exit For
            End If
        Next i
    End If
    
    'Delay at the end of the race
     If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:04")) 'Delay

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RacePresentation()")
    Call basAuxiliary.CodeCrash
End Sub

'Play the kidnapping sequence
Private Sub PlayKidnappingSequence(i As Integer)
    For j = 1 To 100
        Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
            Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 12)) _
            .Interior.Color = vbGreen
        Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
            Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 12)) _
            .Interior.Color = objRace.TRACK_COLOUR
    Next j
    g_arr_varHorses(i, 20) = "KIDNAPPED"
    If objOption.MOMENTUM_BARS Or objOption.MOMENTUM_ICONS _
        Then Cells(g_arr_varHorses(i, 3), 13).ClearContents 'Delete speed value
    With Cells(g_arr_varHorses(i, 3), 12)
        .Value = Cells(g_arr_varHorses(i, 3), 12).Value & " >>> " & GetText(g_arr_Text, "RACESPEC014") '" >>> kidnapped"
        .Interior.Color = vbGreen
    End With
End Sub

'Start of the race
Private Sub RunRace(Optional ml As Boolean, Optional test As Boolean, Optional replay As Boolean)

    'Variable used for statistical purposes
    Dim lngLoop As Long
    
    'Variables used for race information
    Dim strMetres As String
    Dim strLeader As String
    Dim dblProgressBar As Double
    
    'Variables used for water splashes
    Dim intSquirtPattern As Integer
    Dim intSquirtLength As Integer
    Dim dblSquirtColour As Double
    
    'Variable used for air quality
    Dim lngAirPattern As Long
    
    'Variables for overriding settings
    Dim noRefuse As Boolean
    Dim noTactics As Boolean
    Dim noSlipstream As Boolean
    Dim noFavourite As Boolean
    Dim yesAliens As Boolean
    Dim yesTactics As Boolean
    
    'Variables for Race Speed Monitor
    Dim intRSMonRefresh As Integer
    Dim intChartMin As Integer, intChartMax As Integer
    Dim arrSpeedEmpty As Variant 'For the 'zero line' to set the x axis length
    Dim arrSpeedLine1 As Variant, arrSpeedLine2 As Variant, arrSpeedLine3 As Variant
    Dim dblSumSpeed1 As Double, dblSumSpeed2 As Double, dblSumSpeed3 As Double
    Dim lngColSpeedLine1 As Long, lngColSpeedLine2 As Long, lngColSpeedLine3 As Long 'Colour for the speed lines
    Dim arrTempSpeedLogLength() As Double 'Auxiliary array for cutting the speed log array to the perfect length

    On Error GoTo ERRORHANDLING 'In case an error occurs

    Call basAuxiliary.ActivateRaceSheet
    
    'Resize the arrays for the race results
    ReDim m_arr_varPhotofinish(1 To objRace.NUMBER_ENROLLED, 0 To 4) 'Snapshot of the finish
    ReDim m_arr_varResults(0 To objRace.NUMBER_STARTING, 0 To 7) 'Ranking list
    
    If Not replay Then Call StoreRaceReplay1
    
    'Set up the Race Speed Monitor (if RSMon is activated)
    If Not ml And objOption.SPEEDMONITOR Then

        'Set the speed chart refresh interval
        intRSMonRefresh = objRace.METRES / 1000 * objOption.SPEEDMON_REFRESHRATE

        'Set the y axis range
        If objOption.RSMON_DISTANCE Then 'Display the cumulated metres run
            intChartMin = 0
            intChartMax = objRace.METRES
        Else 'Display the current speed
            Select Case objRace.TRACK_SURFACE
                Case "MOON" 'Moon
                    intChartMin = -220
                    intChartMax = 700
                Case "MARS" 'Mars
                    intChartMin = -150
                    intChartMax = 600
                Case "JUPITER" 'Jupiter
                    intChartMin = 50
                    intChartMax = 250
                Case "PLUTO" 'Pluto
                    intChartMin = 200
                    intChartMax = 1700
                Case "SATURN" 'Saturn
                    intChartMin = -600
                    intChartMax = 1000
                Case Else 'All tracks on planet earth
                    intChartMin = (1200 + 10 * objOption.SPEEDMON_REFRESHRATE / 4) * objOption.SPEED_FACTOR
                    intChartMax = (1800 - 10 * objOption.SPEEDMON_REFRESHRATE / 4) * objOption.SPEED_FACTOR
            End Select
        End If
        
        'Initialize the speed lines
        arrSpeedLine1 = Array(0)
        arrSpeedLine2 = Array(0)
        arrSpeedLine3 = Array(0)
    
        'Fill the zero line with zeros
        arrSpeedEmpty = Array()
        ReDim Preserve arrSpeedEmpty(1 To objRace.METRES / intRSMonRefresh * 0.7 / objOption.SPEED_FACTOR) 'Perfect length (estimated)
        For i = 1 To UBound(arrSpeedEmpty)
            arrSpeedEmpty(i) = 0
        Next i
        
        'Draw the chart
        Set m_shSpeedChart = g_wksRace.Shapes.AddChart(xlLine, 2, _
            rows(objBasicData.TOP_ROWS + 2 * objRace.NUMBER_ENROLLED + 7).top + 5, _
                Columns(14).left - 6, _
                (rows(ActiveWindow.VisibleRange.rows.count).top) - (rows(objBasicData.TOP_ROWS + 2 * objRace.NUMBER_ENROLLED + 7).top) - 2)
                    'ChartType, Left, Top, Width, Height
        With m_shSpeedChart
            .name = "SpeedChart"
            With .Line
                .Visible = msoTrue
                .Weight = 1
                .ForeColor.RGB = RGB(0, 0, 0) 'Black border line
            End With
            With .Chart
                .Axes(xlValue).MinimumScale = intChartMin
                .Axes(xlValue).MaximumScale = intChartMax
                If objOption.RSMON_SPEED Then .Axes(xlValue).Delete 'Delete the vertical axes values
                .Axes(xlValue).MajorGridlines.Delete
                .Axes(xlCategory).Delete
                .SeriesCollection.NewSeries
            End With
            
            'Set the x axis by 'drawing' the zero line
            With .Chart.SeriesCollection(1)
                .Values = arrSpeedEmpty
            End With

            'Create a speed line for each selected horse
            For i = 1 To g_colRSMon.count
                .Chart.SeriesCollection.NewSeries
            Next i
            
            'Speed chart formattings
            .Chart.SeriesCollection(1).Format.Line.Visible = False 'Set the 'zero line' invisible
            .Chart.Legend.LegendEntries(1).Delete 'Delete the legend of the 'zero line'
            .Chart.Legend.Position = xlLeft 'Position of the legend
            .Chart.Legend.Font.size = m_intFontSize - g_colRSMon.count
            .Chart.Legend.Font.Bold = True
            .Chart.PlotArea.Format.Fill.TwoColorGradient msoGradientHorizontal, 1
            'Background: Track colour
            .Chart.ChartArea.Format.Fill.ForeColor.RGB = RGB(GetRed(objRace.TRACK_COLOUR), GetGreen(objRace.TRACK_COLOUR), GetBlue(objRace.TRACK_COLOUR))
            'Plot area: Track colour (50% lighter)
            .Chart.PlotArea.Format.Fill.ForeColor.RGB = RGB(GetRed(objRace.TRACK_COLOUR) + (255 - GetRed(objRace.TRACK_COLOUR)) * 50 / 100, _
                            GetGreen(objRace.TRACK_COLOUR) + (255 - GetGreen(objRace.TRACK_COLOUR)) * 50 / 100, _
                            GetBlue(objRace.TRACK_COLOUR) + (255 - GetBlue(objRace.TRACK_COLOUR)) * 50 / 100)
            'Plot area: Track colour (50% darker)
            .Chart.PlotArea.Format.Fill.BackColor.RGB = RGB(GetRed(objRace.TRACK_COLOUR) * (100 - 50) / 100, _
                            GetGreen(objRace.TRACK_COLOUR) * (100 - 50) / 100, _
                            GetBlue(objRace.TRACK_COLOUR) * (100 - 50) / 100)
            'Plot area: Border around
            .Chart.PlotArea.Format.Line.ForeColor.RGB = RGB(0, 0, 0)

            'Prepare the speed line(s) for the selected horse(s)
            For i = 1 To UBound(g_arr_varHorses)
                'First selected horse
                If CDbl(g_arr_varHorses(i, 11)) = CDbl(g_colRSMon(1)) Then
                    'Get the line colour
                    If IsArray(g_arr_varHorses(i, 2)) Then
                        lngColSpeedLine1 = g_arr_varHorses(i, 2)(7)
                    Else
                        lngColSpeedLine1 = g_arr_varHorses(i, 2)
                    End If
                    'Speed line formattings
                    With .Chart.SeriesCollection(2)
                        .name = g_arr_varHorses(i, 1)
                        .Smooth = True
                        .Format.Line.Weight = (m_intFontSize * 0.9) - (2 * g_colRSMon.count)
                        .Format.Line.ForeColor.RGB = RGB(GetRed(lngColSpeedLine1), GetGreen(lngColSpeedLine1), GetBlue(lngColSpeedLine1))
                    End With
                End If
  
                'Second selected horse
                If g_colRSMon.count > 1 Then
                    If CDbl(g_arr_varHorses(i, 11)) = CDbl(g_colRSMon(2)) Then
                        'Get the line colour
                        If IsArray(g_arr_varHorses(i, 2)) Then
                            lngColSpeedLine2 = g_arr_varHorses(i, 2)(7)
                        Else
                            lngColSpeedLine2 = g_arr_varHorses(i, 2)
                        End If
                        'Speed line formattings
                        With .Chart.SeriesCollection(3)
                            .name = g_arr_varHorses(i, 1)
                            .Smooth = True
                            .Format.Line.Weight = (m_intFontSize * 0.9) - (2 * g_colRSMon.count)
                            .Format.Line.ForeColor.RGB = RGB(GetRed(lngColSpeedLine2), GetGreen(lngColSpeedLine2), GetBlue(lngColSpeedLine2))
                        End With
                    End If
                End If
            
                'Third selected horse
                If g_colRSMon.count > 2 Then
                    If CDbl(g_arr_varHorses(i, 11)) = CDbl(g_colRSMon(3)) Then
                        'Get the line colour
                        If IsArray(g_arr_varHorses(i, 2)) Then
                            lngColSpeedLine3 = g_arr_varHorses(i, 2)(7)
                        Else
                            lngColSpeedLine3 = g_arr_varHorses(i, 2)
                        End If
                        'Speed line formattings
                        With .Chart.SeriesCollection(4)
                            .name = g_arr_varHorses(i, 1)
                            .Smooth = True
                            .Format.Line.Weight = (m_intFontSize * 0.9) - (2 * g_colRSMon.count)
                            .Format.Line.ForeColor.RGB = RGB(GetRed(lngColSpeedLine3), GetGreen(lngColSpeedLine3), GetBlue(lngColSpeedLine3))
                        End With
                    End If
                End If
            Next i
            
        End With
        DoEvents
    End If
    
    'Override settings dependent on the selected race
    Select Case objRace.RACE_ID
        Case "SPACE"
            noRefuse = True
            noTactics = True
            noSlipstream = True
            noFavourite = True
            yesAliens = True
        Case "GDIG" 'Gold Diggers 1948
            noRefuse = True
            yesTactics = True 'Always with (special) tactics
        Case "CORONA2020"
            noRefuse = True
            noSlipstream = True
        'Only for testing purposes (adapt the variables)
        Case "TEST" 'Test race
'            yesTactics = True 'Always with (special) tactics
    End Select
    
    'Set the number of running horses equal to the number of starters
    m_intHorsesRunning = objRace.NUMBER_STARTING
    
    'Set the current horse status (RUNNING, REFUSED...)
    If Not replay Then
        Dim intRefuse As Integer
        For i = 1 To UBound(g_arr_varHorses)
            If objOption.REFUSE_RUN And Not noRefuse Then 'Horses can refuse to run (if activated in the race options)
                If g_arr_varHorses(i, 0) = "START" Then
                    Randomize
                    intRefuse = Int(((objOption.REFUSAL_RATE - 1) - 0 + 1) * Rnd + 0) 'Random number between 0 and the value determined in the settings
                    If intRefuse = 0 Then
                        g_arr_varHorses(i, 20) = "REFUSED"
                        m_intHorsesRunning = m_intHorsesRunning - 1
                    Else
                        g_arr_varHorses(i, 20) = "RUNNING"
                    End If
                End If
            ElseIf g_arr_varHorses(i, 0) = "START" Then
                g_arr_varHorses(i, 20) = "RUNNING"
            End If
        Next i
    End If
    
    If replay Then
        For i = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(i, 20) = "REFUSED" Then m_intHorsesRunning = m_intHorsesRunning - 1
        Next i
    End If

    If Not ml Then 'Skip for Machine Learning simulation races
        'Check the air quality
        If objRace.SPECIAL = "PARTICULATES" Then
            lngAirPattern = objOption.PARTICULATES_PATTERN
        Else
            lngAirPattern = xlSolid 'Clean air
        End If
        
        'Prepare the race information on the worksheet (if selected in the race options)
        If objOption.RACE_INFO And objOption.RACE_INFO_WKS Then
    
            'Calculation of the progress bar width
            dblProgressBar = Columns(12).width - 2
            'Formattings
            Call basAuxiliary.RaceInfoWorksheet(objOption.RACE_INFO_COL_B, objOption.RACE_INFO_COL_F, objBasicData.TOP_ROWS, True)
        End If
            
        'Starting procedure with gates
        If objRace.STARTING_GATE = "Y" Then
            'Hide gate numbers
            Range(Cells(7 + objBasicData.TOP_ROWS, objBasicData.LEFT_COLS + 4), Cells(5 + 2 * objRace.NUMBER_ENROLLED + objBasicData.TOP_ROWS, objBasicData.LEFT_COLS + 4)).Value = ""
            If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:04")) 'Delay
            'Open the gates
            Range(Cells(6 + objBasicData.TOP_ROWS, objBasicData.LEFT_COLS + 13), Cells(objRace.NUMBER_ENROLLED * 2 + 6 + objBasicData.TOP_ROWS, objBasicData.LEFT_COLS + 13)).Interior.Color = objRace.TRACK_COLOUR
        End If
            
        If objOption.SPEECH Then Call SpeechOut(GetText(g_arr_Text, "RACE034")) 'Voice output if selected
                
        'Race information data
        strMetres = GetText(g_arr_Text, "RACE008") '"m"
        strLeader = GetText(g_arr_Text, "RACEINFO001") '"The current leader is"
        objStat.LEADER_POSITION = g_arr_varHorses(1, 4) - (objBasicData.LEFT_COLS + 12) 'Position zero
        objStat.LEADER_NAME = ""
    End If
    
    'Reset variables used for the finish
    m_blnPhotofinish = False
    m_blnWin = False
    m_blnDeadHeat = False
    m_intPlace = 1
    m_intFinishLoop = 0
    
    #If Debugging Then 'For debugging purposes: Race start time
        Dim timeStart As Date 'For calculating the race time
        timeStart = Now
        Debug.Print vbNewLine & objRace.RACE_NAME & " (" & objRace.METRES & "m)"
        Debug.Print "RACE START : " & Format(timeStart, "HH:MM:SS") & vbNewLine
    #End If

    If Not ml Then 'Skip for Machine Learning simulation races
        'In case of displaying speed bars for each horse: Show the caption
        If objOption.MOMENTUM_BARS Or objOption.MOMENTUM_ICONS _
            Then Cells(5 + objBasicData.TOP_ROWS, 13).Font.Color = objOption.COL_TEXT
        
        'In case of a tactical race with displaying the section speed: Show the caption
        If objOption.TACTICS And objOption.TACTICS_REVEAL_TAC Then
            Columns(10).Font.Color = objOption.COL_TEXT ' "Reveal" the text
        End If
    End If
        
    #If Debugging Then
        Debug.Print "Speed log array length (initial) : " & UBound(g_arr_varHorses(1, 14))
    #End If
    
    'Game loop for the race
    Do Until m_intPlace > m_intHorsesRunning 'As long as at least one horse is running

        lngLoop = lngLoop + 1
        
        #If Debugging Then
            Debug.Print "Race loop starting: " & lngLoop
        #End If
    
        Call basAuxiliary.ActivateRaceSheet

        'Reset the counter for horses that crossed the finish line in this loop
        m_intHorsesFinishing = 0
            
        'Re-calculation of each horse�s position
        For i = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(i, 20) = "RUNNING" Then 'Only for horses that are still running


            'Enlarge the race speed log array if necessary
            If lngLoop > UBound(g_arr_varHorses(i, 14)) Then
                arrTempSpeedLogLength = g_arr_varHorses(i, 14)
                ReDim Preserve arrTempSpeedLogLength(1 To lngLoop + 99) 'Enlarge length by 100
                g_arr_varHorses(i, 14) = arrTempSpeedLogLength
                #If Debugging Then
                    Debug.Print "Speed log array length (#" & g_arr_varHorses(i, 11) _
                        & " " & g_arr_varHorses(i, 1) & ") enlarged: " & UBound(g_arr_varHorses(i, 14))
                #End If
            End If
            
            'Speed factor for this loop
            g_arr_varHorses(i, 7) = SpeedLoop()

            'For development purposes: Equalise the speed factors for all horses
            'with the sliders in the Developer Tools pop-up
            If g_arr_Developer(1) = 1 Then g_arr_varHorses(i, 5) = 1500 'Basic speed
            If g_arr_Developer(2) = 1 Then g_arr_varHorses(i, 6) = 1500 'Form
            If g_arr_Developer(3) = 1 Then g_arr_varHorses(i, 7) = 1500 'Loop factor
            
            '...or by commenting in the next three lines
'                g_arr_varHorses(i, 5) = 1500 'Basic speed
'                g_arr_varHorses(i, 6) = 1500 'Form
'                g_arr_varHorses(i, 7) = 1500 'Loop factor
            
            #If Debugging Then
                Debug.Print
                Debug.Print "RACE LOOP --> " & lngLoop
                Debug.Print "#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1)
                Debug.Print "BASIC SPEED   " & g_arr_varHorses(i, 5)
                Debug.Print "FORM          " & g_arr_varHorses(i, 6)
                Debug.Print " >AVG BAS/FRM " & (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6)) / 2
                Debug.Print "LOOP          " & g_arr_varHorses(i, 7)
                Debug.Print " >AVG B/F/L   " & (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + g_arr_varHorses(i, 7)) / 3
            #End If

            'Calculate the race section in which the horse runs
            Select Case True
                Case (g_arr_varHorses(i, 4) - objBasicData.LEFT_COLS - 12) < objRace.METRES * 1 / 6 'Section 1/6
                    g_arr_varHorses(i, 10) = 1
                Case (g_arr_varHorses(i, 4) - objBasicData.LEFT_COLS - 12) < objRace.METRES * 2 / 6 'Section 2/6
                    g_arr_varHorses(i, 10) = 2
                Case (g_arr_varHorses(i, 4) - objBasicData.LEFT_COLS - 12) < objRace.METRES * 3 / 6 'Section 3/6
                    g_arr_varHorses(i, 10) = 3
                Case (g_arr_varHorses(i, 4) - objBasicData.LEFT_COLS - 12) < objRace.METRES * 4 / 6 'Section 4/6
                    g_arr_varHorses(i, 10) = 4
                Case (g_arr_varHorses(i, 4) - objBasicData.LEFT_COLS - 12) < objRace.METRES * 5 / 6 'Section 5/6
                    g_arr_varHorses(i, 10) = 5
                Case Else 'Section 6/6
                    g_arr_varHorses(i, 10) = 6
            End Select
                               
            'Calculate the exact step width in this loop
            '-------------------------------------------
            'In case of special tactics (numeric, e.g. 000222 in Gold Diggers Race)
            If (objOption.TACTICS = True And Not noTactics And IsNumeric(Mid(g_arr_varHorses(i, 25), g_arr_varHorses(i, 10), 1))) _
                Or (yesTactics = True And IsNumeric(Mid(g_arr_varHorses(i, 25), g_arr_varHorses(i, 10), 1))) Then
                If objStat.LEADER_POSITION < 2 * objOption.SPEED_FACTOR Then objStat.LEADER_POSITION = 2 * objOption.SPEED_FACTOR
                                 
                'The calculation for horse with special (numeric) tactics
                'is based on the leader�s position
                Dim intSectorLeader As Integer 'Sector in which the leader runs
                intSectorLeader = WorksheetFunction.RoundDown((objStat.LEADER_POSITION - 2 * objOption.SPEED_FACTOR) / (objRace.METRES / 6) + 1, 0)
                If intSectorLeader > 6 Then intSectorLeader = 6 'The calculation exceeds in some cases the value 6 at the end of the race
                g_arr_varHorses(i, 8) = Mid(g_arr_varHorses(i, 25), intSectorLeader, 1) _
                    * Round(((g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + g_arr_varHorses(i, 7)) / 3))

                #If Debugging Then
                    Debug.Print "SPEC TACTICS *" & Mid(g_arr_varHorses(i, 25), intSectorLeader, 1)
                    Debug.Print "          >>> " & g_arr_varHorses(i, 8)
                #End If
                
                If Not ml Then 'Skip for Machine Learning simulation races
                    If objOption.TACTICS_REVEAL_CURR Then
                        Cells(g_arr_varHorses(i, 3), 10).Value = Mid(g_arr_varHorses(i, 25), intSectorLeader, 1) & "x"
                    End If
                End If
                
            'In case of a tactical race (non-numeric, e.g. SMFSMF)
            ElseIf (objOption.TACTICS = True And Not noTactics) Or yesTactics Then
                g_arr_varHorses(i, 8) = _
                    (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                        g_arr_varHorses(i, 7) + g_arr_varHorses(i, 25 + g_arr_varHorses(i, 10))) / 4

                #If Debugging Then
                    Debug.Print "TACTICS       " & g_arr_varHorses(i, 25 + g_arr_varHorses(i, 10))
                    Debug.Print " >AVG B/F/L/T " & g_arr_varHorses(i, 8)
                #End If

                If Not ml Then 'Skip for Machine Learning simulation races
                    If objOption.TACTICS_REVEAL_CURR Then
                        Cells(g_arr_varHorses(i, 3), 10).Value = Mid(g_arr_varHorses(i, 25), g_arr_varHorses(i, 10), 1)
                    End If
                End If

            'In case of no tactics selected in the race options
            Else
                g_arr_varHorses(i, 8) = _
                    (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + g_arr_varHorses(i, 7)) / 3
                #If Debugging Then
                    Debug.Print "NO TACTICS    " & g_arr_varHorses(i, 8)
                #End If
            End If

            If Not ml Then 'Skip for Machine Learning simulation races
                'If chosen: Highlight the current tactic section
                If objOption.TACTICS And objOption.TACTICS_REVEAL_TAC Then
                    Cells(g_arr_varHorses(i, 3), 1) _
                        .Characters(Start:=g_arr_varHorses(i, 10) - 1, Length:=1).Font.Color = vbBlack
                    Cells(g_arr_varHorses(i, 3), 1) _
                        .Characters(Start:=g_arr_varHorses(i, 10), Length:=1).Font.Color = vbYellow
                End If

                'Remove water splashes
                If objRace.SQUIRT = True Then
                        Range(Cells(g_arr_varHorses(i, 3) - 1, g_arr_varHorses(i, 4) - 6), _
                            Cells(g_arr_varHorses(i, 3) + 1, g_arr_varHorses(i, 4) - 14)).Interior.Pattern = xlSolid
                End If
                        
                'Remove slipstream illustration
                If objOption.SLIPSTREAM_IMPACT > 0 And objOption.SLIPSTREAM_SHOW And g_arr_varHorses(i, 22) > 0 _
                    And Not noSlipstream Then
                        If g_arr_varHorses(i, 4) <= objRace.METRES + objBasicData.LEFT_COLS + 9 Then
                            Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9), _
                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 12)).Interior.Pattern = lngAirPattern
                        End If
                End If
            End If

            'Reset slipstream value
            g_arr_varHorses(i, 22) = 0
            
            'Calculate slipstream value (if activated in the race options)
            If objOption.SLIPSTREAM_IMPACT > 0 And Not noSlipstream Then
                For k = 1 To UBound(g_arr_varHorses) 'Loop through the horses
                    'Find an adjacent horse
                    If g_arr_varHorses(i, 15) - 1 = g_arr_varHorses(k, 15) _
                        Or g_arr_varHorses(i, 15) + 1 = g_arr_varHorses(k, 15) Then 'One row above or below
                            'Check the distance to the horse and decide whether slipstream is given
                            If g_arr_varHorses(i, 4) > g_arr_varHorses(k, 4) - 8 _
                                And g_arr_varHorses(i, 4) < g_arr_varHorses(k, 4) - 4 Then
                                    'Determine the multiplication factor
                                    If objOption.SLIPSTREAM_IMPACT = 2 Then
                                        g_arr_varHorses(i, 22) = g_arr_varHorses(i, 22) + 1
                                    Else
                                        g_arr_varHorses(i, 22) = 1
                                    End If
                                    #If Debugging Then
                                        Debug.Print "IN SLIPSTR OF #" & g_arr_varHorses(k, 11) & " " _
                                            & g_arr_varHorses(k, 1) & " (FACTOR " & g_arr_varHorses(i, 22) & ")"
                                    #End If
                            End If
                    End If
                Next k
            End If

            If Not replay Then
                'Take the slipstream effect into account
                g_arr_varHorses(i, 8) = g_arr_varHorses(i, 8) + (g_arr_varHorses(i, 22) * 100)
                
                #If Debugging Then
                    If objOption.SLIPSTREAM_IMPACT > 0 And g_arr_varHorses(i, 22) > 0 _
                        And Not noSlipstream Then
                            Debug.Print "SLIPSTREAM   +" & g_arr_varHorses(i, 22) * 100
                            Debug.Print "          >>> " & g_arr_varHorses(i, 8)
                    End If
                #End If
    
                'Multiply position with the race speed factor
                g_arr_varHorses(i, 8) = g_arr_varHorses(i, 8) * objOption.SPEED_FACTOR
                    
                #If Debugging Then
                    Debug.Print "SPEEDFACTOR  *" & objOption.SPEED_FACTOR
                    Debug.Print "          >>> " & g_arr_varHorses(i, 8)
                    Debug.Print "POSITION OLD >> " & g_arr_varHorses(i, 9)
                    Debug.Print "COLUMN OLD >> " & g_arr_varHorses(i, 4)
                    Debug.Print "STEP WIDTH   >> " & g_arr_varHorses(i, 8)
                    Debug.Print "POSITION NEW >> " & g_arr_varHorses(i, 9) + g_arr_varHorses(i, 8)
                    Debug.Print "COLUMN NEW >> " & objBasicData.LEFT_COLS + 12 + Round((g_arr_varHorses(i, 9) + g_arr_varHorses(i, 8)) / 1000, 0)
                    Debug.Print "METRES RUN >> " & Format(Round((g_arr_varHorses(i, 9) + g_arr_varHorses(i, 8)) / 1000, 2), "0.00")
                #End If

            Else 'Get the step width from the race speed log array
                g_arr_varHorses(i, 8) = g_arr_varHorses(i, 14)(lngLoop)
            End If

                If Not ml Then 'Skip for Machine Learning simulation races
                    'Calculate and display the momentum
                    If objOption.MOMENTUM_BARS Or objOption.MOMENTUM_ICONS Then
                        If g_arr_varHorses(i, 20) = "RUNNING" Then
                            'Insert the latest speed value into the array
                            g_arr_varHorses(i, 19)(lngLoop Mod objOption.MOMENTUM_REFRESHRATE + 1) = g_arr_varHorses(i, 8)
                            'Refresh the current speed dependent by taking the average of the latest speed values
                            Cells(g_arr_varHorses(i, 3), 13).Value = WorksheetFunction.Average(g_arr_varHorses(i, 19))
                        End If
                    End If
                End If
            End If
        Next i 'End of the re-calculation of the position

        #If Debugging Then
            Debug.Print 'Blank line
        #End If
            
        'Horses are running
        For i = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(i, 20) = "RUNNING" Then 'Only for horses that are still running
            
                If Not ml Then 'Skip for Machine Learning simulation races
                    Call basAuxiliary.ActivateRaceSheet
                                            
                    'Delete the horse on the worksheet
                    'by assigning the track colour to the horse�s position
                    Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
                        Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)) _
                        .Interior.Color = objRace.TRACK_COLOUR

                If Not replay Then
                    'In case of a aliens around: Check the alien behaviour
                    If yesAliens Then
                        If objOption.SPACE_ALIENS = enumAliens.unfriendly Then
                            Dim lngKidnapping As Long
                            Randomize
                            lngKidnapping = Int(((20000 / objOption.SPACE_KIDNAPPINGRATE) - 0 + 1) * Rnd + 0) 'Random number between 0 and 20000
                            If lngKidnapping = 0 Then
                                Call PlayKidnappingSequence(i)
                                m_intHorsesRunning = m_intHorsesRunning - 1
                                g_arr_varHorses(i, 14)(lngLoop) = -999999 'Assign a number that represents kidnapping
                                Exit For
                            End If
                        End If
                    End If
                Else 'Check for kidnapping in a replay
                    If g_arr_varHorses(i, 14)(lngLoop) = -999999 Then
                        Call PlayKidnappingSequence(i)
                        m_intHorsesRunning = m_intHorsesRunning - 1
                        Exit For
                    End If
                End If
            End If

            'Calculate the new position of the horse
            g_arr_varHorses(i, 9) = g_arr_varHorses(i, 9) + g_arr_varHorses(i, 8) 'Exact (internal) position (accuracy 0.25 millimetres)

            If Not replay Then
                'Write the current speed to the race speed log
                g_arr_varHorses(i, 14)(lngLoop) = g_arr_varHorses(i, 8)
            End If
            
                'Make sure that no horse moves behind the starting line
                If g_arr_varHorses(i, 9) < 0 Then g_arr_varHorses(i, 9) = 0
                'Calculate the position on the worksheet (accuracy 1 metre)
                g_arr_varHorses(i, 4) = objBasicData.LEFT_COLS + 12 + Round(g_arr_varHorses(i, 9) / 1000, 0)
                        
                If Not ml Then 'Skip for Machine Learning simulation races
                    'Draw the speed line(s) if selected in the race options
                    If objOption.SPEEDMONITOR Then
                        If lngLoop Mod intRSMonRefresh = 0 Then
                            'First selected horse
                            If g_arr_varHorses(i, 11) = CInt(g_colRSMon(1)) Then
                                dblSumSpeed1 = 0
                                For k = lngLoop + 1 - intRSMonRefresh To lngLoop
                                    dblSumSpeed1 = dblSumSpeed1 + g_arr_varHorses(i, 14)(k)
                                Next
                                If (lngLoop / intRSMonRefresh) > UBound(arrSpeedLine1) Then
                                    ReDim Preserve arrSpeedLine1(1 To UBound(arrSpeedLine1) + 1)
                                End If
                                    If objOption.RSMON_DISTANCE Then
                                        arrSpeedLine1(lngLoop / intRSMonRefresh) = g_arr_varHorses(i, 9) / 1000
                                    Else
                                        arrSpeedLine1(lngLoop / intRSMonRefresh) = dblSumSpeed1 / intRSMonRefresh
                                    End If
                                With m_shSpeedChart
                                    With .Chart.SeriesCollection(2)
                                        .Values = arrSpeedLine1
                                    End With
                                End With
                                If g_colRSMon.count = 1 Then DoEvents
                            End If
                            'Second selected horse
                            If g_colRSMon.count > 1 Then
                                If g_arr_varHorses(i, 11) = CInt(g_colRSMon(2)) Then
                                    dblSumSpeed2 = 0
                                    For k = lngLoop + 1 - intRSMonRefresh To lngLoop
                                        dblSumSpeed2 = dblSumSpeed2 + g_arr_varHorses(i, 14)(k)
                                    Next
                                    If (lngLoop / intRSMonRefresh) > UBound(arrSpeedLine2) Then
                                        ReDim Preserve arrSpeedLine2(1 To UBound(arrSpeedLine2) + 1)
                                    End If
                                    If objOption.RSMON_DISTANCE Then
                                        arrSpeedLine2(lngLoop / intRSMonRefresh) = g_arr_varHorses(i, 9) / 1000
                                    Else
                                        arrSpeedLine2(lngLoop / intRSMonRefresh) = dblSumSpeed2 / intRSMonRefresh
                                    End If
                                    With m_shSpeedChart
                                        With .Chart.SeriesCollection(3)
                                            .Values = arrSpeedLine2
                                        End With
                                    End With
                                    If g_colRSMon.count = 2 Then DoEvents
                                End If
                            End If
                            'Third selected horse
                            If g_colRSMon.count > 2 Then
                                If g_arr_varHorses(i, 11) = CInt(g_colRSMon(3)) Then
                                    dblSumSpeed3 = 0
                                    For k = lngLoop + 1 - intRSMonRefresh To lngLoop
                                        dblSumSpeed3 = dblSumSpeed3 + g_arr_varHorses(i, 14)(k)
                                    Next
                                    If (lngLoop / intRSMonRefresh) > UBound(arrSpeedLine3) Then
                                        ReDim Preserve arrSpeedLine3(1 To UBound(arrSpeedLine3) + 1)
                                    End If
                                    If objOption.RSMON_DISTANCE Then
                                        arrSpeedLine3(lngLoop / intRSMonRefresh) = g_arr_varHorses(i, 9) / 1000
                                    Else
                                        arrSpeedLine3(lngLoop / intRSMonRefresh) = dblSumSpeed3 / intRSMonRefresh
                                    End If
                                    With m_shSpeedChart
                                        With .Chart.SeriesCollection(4)
                                            .Values = arrSpeedLine3
                                        End With
                                    End With
                                    DoEvents
                                End If
                            End If
                            DoEvents
                        End If
                    End If
                    
                    'Draw the slipstream illustration (if selected in the race options)
                    If objOption.SLIPSTREAM_IMPACT > 0 And objOption.SLIPSTREAM_SHOW And g_arr_varHorses(i, 22) > 0 _
                        And g_arr_varHorses(i, 4) > objBasicData.LEFT_COLS + 12 _
                        And g_arr_varHorses(i, 4) <= objRace.METRES + objBasicData.LEFT_COLS + 7 _
                        And Not noSlipstream Then
                            With Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9), _
                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 12))
                                If g_arr_varHorses(i, 22) = 1 Then 'Single slipstream effect: slim horizontal lines
                                    .Interior.Pattern = xlLightHorizontal
                                Else 'Double slipstream effect: thick horizontal lines
                                    .Interior.Pattern = xlHorizontal
                                End If
                            End With
                        #If Debugging Then 'For debugging purposes
                            If g_arr_varHorses(i, 22) = 1 Then 'Single slipstream effect: light blue
                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9).Interior.Color = 14395790
                            Else 'Double slipstream effect: dark blue
                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9).Interior.Color = 9851952
                            End If
                        #End If
                    End If
                            
                    'Draw water splashes
                    If objRace.SQUIRT = True And g_arr_varHorses(i, 4) > 13 + objBasicData.LEFT_COLS _
                        And g_arr_varHorses(i, 4) <= objRace.METRES + objBasicData.LEFT_COLS Then
                        
                        If Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)).Interior.Color = objOption.PUDDLE_COL Then
                        
                            'Calculate the squirt pattern by random
                            Randomize 'Initialize the random number generator
                            intSquirtPattern = Int((18 - 16 + 1) * Rnd + 16) 'Value between 16 and 18
                            intSquirtLength = Int((8 - 4 + 1) * Rnd + 4) 'Value between 4 and 8
                            dblSquirtColour = (Int(2 - 0 + 0) * Rnd + 0) - 1 'Value between -1.000... and +1.000...
        
                            With Range(Cells(g_arr_varHorses(i, 3) - 1, g_arr_varHorses(i, 4) - 6), _
                                Cells(g_arr_varHorses(i, 3) + 1, g_arr_varHorses(i, 4) - 6 - intSquirtLength)).Interior 'Squirt length between 4 and 8 metres
                                    .Pattern = intSquirtPattern 'One out of three patterns: 16=xlCrissCross 17=xlGray25 18=xlGray8
                                    .PatternThemeColor = xlThemeColorDark1
                                    'Assign one out of 21 shades by rounding from Double to a value between -1.0 and +1.0
                                    .PatternTintAndShade = Round(dblSquirtColour, 1) 'Shade of the theme colour
                                    'Rounding to 2 digits (dblSquirtColour, 2) leads to 201 different shades,
                                    'however the more shades the slower is the rendering.
                                    'The absolute maximum of different cell formats in a workbook is approx. 64000
                            End With
                        End If
                    End If
                            
                    'Re-paint the horses
                    Call PaintHorse(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7, g_arr_varHorses(i, 2))
        
                    'In case of a mudflats race
                    If objRace.TRACK_SURFACE = "M" Then
                    
                        'Hide the lugworms under the horse by overpainting with the horse colour
                        If IsArray(g_arr_varHorses(i, 2)) Then
                            For j = 0 To 7 'Loop through each segment
                                        'of the horse
                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7 + j) _
                                    .Font.Color = g_arr_varHorses(i, 2)(j) 'Font colour
                            Next j
                        Else
                            Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7), _
                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4))) _
                                .Font.Color = g_arr_varHorses(i, 2) 'Font colour
                        End If
                        
                        'Show the lugworms again behind the horse
                        'by assigning the lugworm colour
                        Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8), _
                            Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - 2 * objOption.SPEED_FACTOR)) _
                                .Font.Color = objOption.LUGWORM_COL 'Font colour
                                
                        'Show a trampled lugworm behind the horse if is hit by a hoof
                        If Not IsEmpty(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8)) _
                                And Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8) <> "|" Then
                                With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8)
                                    .Value = ChrW(1154) 'Cyrillic thousands sign (UTF 1154)
                                End With
                        End If
                        
                        'Restore the track colour behind the horse
                        For k = 1 To 2 * objOption.SPEED_FACTOR 'Take the speed factor into account
                            Select Case Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - k).Value
                                Case "|" 'Indicator for a cell with water
                                    'Overpaint with the colour of puddles
                                    With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - k)
                                        .Font.Color = objOption.PUDDLE_COL 'Font colour
                                        .Interior.Color = objOption.PUDDLE_COL 'Cell colour
                                    End With
                                    #If Debugging Then 'Show the water indicator sign in red
                                        Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - k).Font.Color = vbRed
                                    #End If
                                Case Else
                                    'Overpaint with the original track colour
                                    With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - k)
                                        .Interior.Color = objRace.TRACK_COLOUR 'Cell colour
                                    End With
                            End Select
                        Next k
            
                    End If
                        
                    'SPECIAL: Wild boar track devastation illustration
                    If g_arr_varHorses(i, 24) = "WILD" _
                         And g_arr_varHorses(i, 4) < objRace.METRES + objBasicData.LEFT_COLS + 12 Then
                         'Draw a race track devastation sign (#) under the left segment of the wild boar
                         If IsArray(g_arr_varHorses(i, 2)) Then 'Multicoloured wild boar
                            With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)
                                .Font.Color = g_arr_varHorses(i, 2)(0) 'Colour of the left segment
                                .Value = "#"
                            End With
                        Else 'Monochrome wild boar
                            With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)
                                .Font.Color = g_arr_varHorses(i, 2) 'Wild boar colour
                                .Value = "#"
                            End With
                        End If
                    End If
        
                    'Draw hoof prints behind the horse (if selected in the race options)
                    If objOption.HOOFPRINTS And IsEmpty(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8)) _
                        And g_arr_varHorses(i, 4) > 13 + objBasicData.LEFT_COLS _
                        And Not g_arr_varHorses(i, 24) = "WILD" Then _
                        Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8).Value = "'-" 'Hoof print character
                End If
            End If
            
            If Not ml Then 'Skip for Machine Learning simulation races
                'Horizontal scrolling dependent on the camera mode
                If objOption.FOCUSED_RUN = enumCamera.standard Then 'Scrolling in standard mode
                    'Check whether the leading horse is near the right edge of the screen
                    If g_arr_varHorses(i, 4) > ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column _
                                - ((ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column _
                                - ActiveWindow.VisibleRange.Column) * 1 / 10) _
                        And ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column <= objRace.METRES + objBasicData.LEFT_COLS + 2 Then
                            'Standard scrolling in the style of paging
                            ActiveWindow.ScrollColumn = ActiveWindow.VisibleRange.Column _
                                                + ((ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column _
                                                - ActiveWindow.VisibleRange.Column) * 8 / 10)
                    End If
                Else 'Scrolling in a Focused Run
                    If objOption.FOCUSED_RUN = enumCamera.focus_horse Then 'Find the horse in focus
                        If g_arr_varHorses(i, 11) = objOption.FOCUSED_NR And g_arr_varHorses(i, 20) = "RUNNING" Then
                            'Check whether the focused horse is in the middle of the screen
                            If g_arr_varHorses(i, 4) > ((ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column _
                                    - ActiveWindow.VisibleRange.Column) / 2) + ActiveWindow.VisibleRange.Column Then
                            
                                #If Debugging Then
                                    Debug.Print
                                    Debug.Print "ScrollColumn before: " & ActiveWindow.ScrollColumn
                                    Debug.Print "Scroll to the right: " & (g_arr_varHorses(i, 4) - ActiveWindow.ScrollColumn) - (ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column - g_arr_varHorses(i, 4))
                                #End If
                                
                                'Scroll to the right dependent on the new position of the focused horse
                                ActiveWindow.ScrollColumn = ActiveWindow.ScrollColumn + (g_arr_varHorses(i, 4) - ActiveWindow.ScrollColumn) - (ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column - g_arr_varHorses(i, 4))
                                
                                #If Debugging Then
                                    Debug.Print "ScrollColumn after : " & ActiveWindow.ScrollColumn 'For debugging purposes
                                #End If
                            End If
                        End If
                    Else 'Find the leader
                        If g_arr_varHorses(i, 11) = objStat.LEADER_NR And g_arr_varHorses(i, 20) = "RUNNING" Then
                            'Check whether the focused horse is in the middle of the screen
                            If g_arr_varHorses(i, 4) > ((ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column _
                                    - ActiveWindow.VisibleRange.Column) / 2) + ActiveWindow.VisibleRange.Column Then
                            
                                #If Debugging Then
                                    Debug.Print
                                    Debug.Print "ScrollColumn before: " & ActiveWindow.ScrollColumn
                                    Debug.Print "Scroll to the right: " & (g_arr_varHorses(i, 4) - ActiveWindow.ScrollColumn) - (ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column - g_arr_varHorses(i, 4))
                                #End If
                                
                                'Scroll to the right dependent on the new position of the focused horse
                                ActiveWindow.ScrollColumn = ActiveWindow.ScrollColumn + (g_arr_varHorses(i, 4) - ActiveWindow.ScrollColumn) - (ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column - g_arr_varHorses(i, 4))
                                
                                #If Debugging Then
                                    Debug.Print "ScrollColumn after : " & ActiveWindow.ScrollColumn 'For debugging purposes
                                #End If
                            End If
                        End If
                    End If
                End If
    
                'Refresh the data of the current leader
                If (g_arr_varHorses(i, 4) - objBasicData.LEFT_COLS - 12) > objStat.LEADER_POSITION Then
                    objStat.LEADER_POSITION = g_arr_varHorses(i, 4) - objBasicData.LEFT_COLS - 12 'Position of the leader
                    objStat.LEADER_NAME = g_arr_varHorses(i, 1) 'Name of the leader
                    objStat.LEADER_NR = g_arr_varHorses(i, 11) 'Number of the leader
                End If
            End If
        Next i
        
        If Not ml Then 'Skip for Machine Learning simulation races
            'Show race information (if selected in the race options)
            If objOption.RACE_INFO Then
                'Refresh the race information data (on the worksheet)
                If objOption.RACE_INFO_WKS Then
                    'Race distance progress bar
                    If objOption.RACE_INFO_PROGRESS Then
                        With Cells(3 + objBasicData.TOP_ROWS, 11)
                            'Relation of metres run and the total race distance
                            .Value = objStat.LEADER_POSITION & strMetres & " / " & objRace.METRES & strMetres
                        End With
                        
                        g_shpBar.width = dblProgressBar / objRace.METRES * objStat.LEADER_POSITION
                            'Width of the progress bar
                        DoEvents
                    End If
                    'Name of the leader
                    If objOption.RACE_INFO_LEADER Then
                        'Only in the range between 20 meters after the start and 20 meters before the finish
                        If objStat.LEADER_POSITION < 20 Or objStat.LEADER_POSITION > (objRace.METRES - 20) Then
                            Cells(1 + objBasicData.TOP_ROWS, 2).Value = " "
                            Cells(2 + objBasicData.TOP_ROWS, 11).Value = " "
                        'Don�t refresh in every loop
                        ElseIf objStat.LEADER_POSITION Mod 5 = 0 Then
                            With Cells(1 + objBasicData.TOP_ROWS, 2)
                                .Value = strLeader
                            End With
                            With Cells(2 + objBasicData.TOP_ROWS, 11)
                                .Value = objStat.LEADER_NAME
                            End With
                        End If
                    End If
                End If
                
                'Refresh the race information data (in a pop-up)
                If objOption.RACE_INFO_POP Then
                    'Race distance progress bar
                    If objOption.RACE_INFO_PROGRESS Then
                        With frmRaceInfo.Controls("lbl_RI3b_dyn")
                            .width = CInt(200 / objRace.METRES * objStat.LEADER_POSITION)
                            .caption = objStat.LEADER_POSITION
                        End With
                    End If
                    'Name of the leader
                    If objOption.RACE_INFO_LEADER Then
                        'Only in the range between 20 meters after the start and 20 meters before the finish
                        If objStat.LEADER_POSITION < 20 Or objStat.LEADER_POSITION > (objRace.METRES - 20) Then
                            frmRaceInfo.Controls("lbl_RI4a_dyn").caption = " "
                            frmRaceInfo.Controls("lbl_RI4b_dyn").caption = " "
                        'Don�t refresh in every loop
                        ElseIf objStat.LEADER_POSITION Mod 5 = 0 Then
                            frmRaceInfo.Controls("lbl_RI4a_dyn").caption = strLeader
                            frmRaceInfo.Controls("lbl_RI4b_dyn").caption = objStat.LEADER_NAME
                        End If
                    End If
                End If
            End If
        End If

        'Check whether one or more horses have reached the finish line
        For i = 1 To UBound(g_arr_varHorses) 'Loop through the horses
            If g_arr_varHorses(i, 20) = "RUNNING" Then 'Only for horses that are still running
            
                #If Debugging Then 'For debugging purposes: Name and position (accuracy 10cm)
                    Debug.Print "#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) & " - Position: " _
                        & Format(Round(g_arr_varHorses(i, 9) / 1000, 2), "0.00") & " metres"
                #End If
            
                If g_arr_varHorses(i, 9) >= objRace.METRES * 1000 Then 'Horse has crossed the finish line
                    #If Debugging Then
                        Debug.Print "  >> FINISHED - Exact position: " & g_arr_varHorses(i, 9)
                    #End If
                    g_arr_varHorses(i, 20) = "CALCULATION" 'Set the horse�s status
                    m_intHorsesFinishing = m_intHorsesFinishing + 1 'Count the number of horses that pass the finish line in this loop
                    
                    If Not ml Then 'Skip for Machine Learning simulation races
                        If objOption.MOMENTUM_BARS Or objOption.MOMENTUM_ICONS _
                            Then Cells(g_arr_varHorses(i, 3), 13).ClearContents
                        If objOption.TACTICS Then
                            If objOption.TACTICS_REVEAL_CURR Then Cells(g_arr_varHorses(i, 3), 10).ClearContents
                            If objOption.TACTICS_REVEAL_TAC Then Cells(g_arr_varHorses(i, 3), 1).Font.Color = objOption.COL_TEXT
                        End If
                    End If
                
                End If
            End If
        Next i
        
        'Evaluate the number of finishers in this loop
        If m_intHorsesFinishing > 0 Then
            If m_blnWin = False Then 'If there is no winner yet
                If m_intHorsesFinishing > 1 Then
                    m_blnPhotofinish = True 'In case of more than one possible winners: Flag for a tight finish
                End If
                
            If Not ml Then 'Skip for Machine Learning simulation races
                Call CreateFinishPhoto(replay) 'Create a photo of the finish
            End If
            
            End If
            m_blnWin = True 'Set true so that the photo of the finish is only created once
            Call CalculateRanking 'Calculate the ranking
        End If

        DoEvents 'Force rendering on the worksheet
    
    Loop 'End of the race loop
    
    #If Debugging Then 'For debugging purposes: Calculate the race time
        Debug.Print vbNewLine & "RACE FINISH: " & Format(Now, "HH:MM:SS")
        Debug.Print "RACE TIME  : " & Format(Now - timeStart, "HH:MM:SS")
        Debug.Print "RACE LOOPS : " & lngLoop & vbNewLine
    #End If
        
    If Not replay Then Call StoreRaceReplay2(lngLoop)
    If objOption.AUTO_SAVE And Not replay Then Call SaveRaceForReplay(True)
        
    #If Debugging Then 'Reveal the minimum and maximum speed across all horses
        If Not ml And objOption.SPEEDMONITOR Then
            Dim mx As Integer, mn As Integer
            mx = 0 'Initial maximum value
            mn = 2000 'Initial minimum value
            For i = 1 To UBound(arrSpeedLine1)
                If arrSpeedLine1(i) > mx Then mx = arrSpeedLine1(i)
                If arrSpeedLine1(i) < mn Then mn = arrSpeedLine1(i)
            Next
            For i = 1 To UBound(arrSpeedLine2)
                If arrSpeedLine2(i) > mx Then mx = arrSpeedLine2(i)
                If arrSpeedLine2(i) < mn Then mn = arrSpeedLine2(i)
            Next
            For i = 1 To UBound(arrSpeedLine3)
                If arrSpeedLine3(i) > mx Then mx = arrSpeedLine3(i)
                If arrSpeedLine3(i) < mn Then mn = arrSpeedLine3(i)
            Next
            Debug.Print "RSMon refresh rate: " & objOption.SPEEDMON_REFRESHRATE
            Debug.Print "   Minimum speed: " & mn
            Debug.Print "   Maximum speed: " & mx
        End If
    #End If
        
        If Not ml Then 'Skip for Machine Learning simulation races
            'Remove race information
            If objOption.RACE_INFO Then
                If objOption.RACE_INFO_POP Then Unload frmRaceInfo 'Close the pop-up
                If objOption.RACE_INFO_WKS Then Call basAuxiliary.RaceInfoWorksheet(objOption.COL_BACK, 0, objBasicData.TOP_ROWS, False) 'Delete on the worksheet
        End If
        
        'Remove section race speed caption
        If objOption.TACTICS_REVEAL_CURR Then Cells(5 + objBasicData.TOP_ROWS, 10).ClearContents
        
        'Remove race speed caption
        If objOption.MOMENTUM_BARS Or objOption.MOMENTUM_ICONS Then Cells(5 + objBasicData.TOP_ROWS, 13).ClearContents
        
        If Not replay Then
            'In case of a photo finish
            If m_blnPhotofinish = True Then
                Call basAuxiliary.ActivateRaceSheet
                If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:04")) 'Delay
                'Clear text "PHOTO FINISH!"
                With Cells(2 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))
                    .Clear
                    .Interior.Color = objOption.COL_BACK
                    If objRace.SPECIAL = "PARTICULATES" Then .Interior.Pattern = objOption.PARTICULATES_PATTERN
                End With
                'Unfreeze the window pane if it is frozen
                If objOption.NAMES_LEFT Or objOption.COLOURS_LEFT Or objOption.HIGHLIGHT_FAV _
                    Or objOption.MOMENTUM_BARS Or objOption.MOMENTUM_ICONS Or objOption.TACTICS_REVEAL_TAC _
                    Or objOption.TACTICS_REVEAL_CURR Or (objOption.FOCUSED_RUN <> enumCamera.standard And objOption.HIGHLIGHT_FOC) _
                    Or (objOption.RACE_INFO And objOption.RACE_INFO_WKS) _
                        Then Call basAuxiliary.Freeze(0, 0, False)
                'Scroll to the area where the photo will be displayed
                ActiveWindow.ScrollColumn = objRace.METRES + objBasicData.LEFT_COLS + 15 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)
                'Black photo background
                Range(Cells(5 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), _
                    Cells(objRace.NUMBER_ENROLLED * 2 + 7 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 175 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR))).Interior.ColorIndex = 1
                If objOption.AUTOFIT Then Call AutoZoom("FinishPhoto")
                'Display text "Photo creation"
                With Cells(4 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))
                    .Value = GetText(g_arr_Text, "RACE026")
                    .Font.Color = objOption.COL_TEXT
                End With
                If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:04")) 'Delay
                'Show the photo of the tight finish
                Call DrawFinishPhoto
                'Display text "Photo evaluation"
                With Cells(4 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))
                    .Value = GetText(g_arr_Text, "RACE027")
                    .Font.Color = objOption.COL_TEXT
                End With
            End If
    
            If test And objOption.SPEEDMONITOR Then m_shSpeedChart.Delete
            If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:02")) 'Delay
            'Clear text above the photo
            Cells(4 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)).ClearContents
        End If
    End If

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RunRace()")
    Call basAuxiliary.CodeCrash
End Sub

'Store the data of the last race for replaying (part 1, before the race)
Private Sub StoreRaceReplay1()

    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    'Prepare the race replay data storage
    ReDim g_arr_varReplay_RaceData(1 To 24, 0 To 1)
    ReDim g_arr_varReplay_HorseData(1 To UBound(g_arr_varHorses), 0 To 31)
    
    Dim x As Integer, y As Integer
    Dim temp() As Variant

    Dim tempDate(1 To 5) As Integer
    tempDate(1) = Year(Now)
    tempDate(2) = Month(Now)
    tempDate(3) = Day(Now)
    tempDate(4) = Hour(Now)
    tempDate(5) = Minute(Now)
    
    'Store the race data
    g_arr_varReplay_RaceData(1, 0) = "GaloppSim Version"
    g_arr_varReplay_RaceData(1, 1) = g_c_version
    g_arr_varReplay_RaceData(2, 0) = "Date (run on)"
    g_arr_varReplay_RaceData(2, 1) = tempDate
    g_arr_varReplay_RaceData(3, 0) = "RACE_ID"
    g_arr_varReplay_RaceData(3, 1) = objRace.RACE_ID
    g_arr_varReplay_RaceData(4, 0) = "PARTICIPANTS"
    g_arr_varReplay_RaceData(4, 1) = objRace.PARTICIPANTS
    g_arr_varReplay_RaceData(5, 0) = "RACE_NAME"
    g_arr_varReplay_RaceData(5, 1) = objRace.RACE_NAME
    g_arr_varReplay_RaceData(6, 0) = "RACE_YEAR"
    g_arr_varReplay_RaceData(6, 1) = objRace.RACE_YEAR
    g_arr_varReplay_RaceData(7, 0) = "TRACK_LOCATION"
    g_arr_varReplay_RaceData(7, 1) = objRace.TRACK_LOCATION
    g_arr_varReplay_RaceData(8, 0) = "COUNTRY_CODE"
    g_arr_varReplay_RaceData(8, 1) = objRace.COUNTRY_CODE
    If objRace.COUNTRY_CODE = "MOON" Then
        g_arr_varReplay_RaceData(24, 0) = "SPACE_PLANET"
        g_arr_varReplay_RaceData(24, 1) = objOption.SPACE_PLANET
    End If
    g_arr_varReplay_RaceData(9, 0) = "TRACK_NAME"
    g_arr_varReplay_RaceData(9, 1) = objRace.TRACK_NAME
    g_arr_varReplay_RaceData(10, 0) = "TRACK_COLOUR"
    g_arr_varReplay_RaceData(10, 1) = objRace.TRACK_COLOUR
    g_arr_varReplay_RaceData(11, 0) = "TRACK_SURFACE"
    g_arr_varReplay_RaceData(11, 1) = objRace.TRACK_SURFACE
    g_arr_varReplay_RaceData(12, 0) = "RACE_TYPE"
    g_arr_varReplay_RaceData(12, 1) = objRace.RACE_TYPE
    g_arr_varReplay_RaceData(13, 0) = "METRES"
    g_arr_varReplay_RaceData(13, 1) = objRace.METRES
    g_arr_varReplay_RaceData(14, 0) = "STARTING_GATE"
    g_arr_varReplay_RaceData(14, 1) = objRace.STARTING_GATE
    g_arr_varReplay_RaceData(15, 0) = "NUMBER_ENROLLED"
    g_arr_varReplay_RaceData(15, 1) = objRace.NUMBER_ENROLLED
    g_arr_varReplay_RaceData(16, 0) = "NUMBER_STARTING"
    g_arr_varReplay_RaceData(16, 1) = objRace.NUMBER_STARTING
    g_arr_varReplay_RaceData(17, 0) = "LANES_FIX_OR_RANDOM"
    g_arr_varReplay_RaceData(17, 1) = objRace.LANES_FIX_OR_RANDOM
    g_arr_varReplay_RaceData(18, 0) = "ADVERTISING"
    g_arr_varReplay_RaceData(18, 1) = objRace.ADVERTISING
    g_arr_varReplay_RaceData(19, 0) = "SPECIAL"
    g_arr_varReplay_RaceData(19, 1) = objRace.SPECIAL

    g_arr_varReplay_RaceData(20, 0) = "ADVERTISEMENT"
    If objRace.ADVERTISING = "Y" Then
        y = GetColumn(g_wksRaceData, "ADVERTISEMENT")
        x = g_wksRaceData.Cells(rows.count, y).End(xlUp).row - 1
        ReDim temp(1 To x)
        For i = 1 To x
            temp(i) = g_wksRaceData.Cells(i + 1, y).Value
        Next i
        g_arr_varReplay_RaceData(20, 1) = temp()
    End If

    g_arr_varReplay_RaceData(21, 0) = "TRIBUNES"
    If objOption.TRIBUNES = True Then
        g_arr_varReplay_RaceData(21, 1) = "Y"
        g_arr_varReplay_RaceData(22, 0) = "TRACK GRAPHICS"
        y = GetColumn(g_wksRaceData, "TRACK GRAPHICS")
        x = g_wksRaceData.Cells(rows.count, y).End(xlUp).row - 1
        ReDim temp(1 To x)
        For i = 1 To x
            temp(i) = g_wksRaceData.Cells(i + 1, y).Value
        Next i
        g_arr_varReplay_RaceData(22, 1) = temp()
    End If
    
    g_arr_varReplay_RaceData(23, 0) = "SPECTATORS"
    g_arr_varReplay_RaceData(23, 1) = objOption.SPECTATORS

    'Store horse data
    For i = 1 To g_arr_varReplay_RaceData(15, 1) 'NUMBER_ENROLLED
        g_arr_varReplay_HorseData(i, 0) = g_arr_varHorses(i, 0) 'Status before the race
        g_arr_varReplay_HorseData(i, 1) = g_arr_varHorses(i, 1) 'Name of the horse
        g_arr_varReplay_HorseData(i, 2) = g_arr_varHorses(i, 2) 'Horse colour
        g_arr_varReplay_HorseData(i, 3) = g_arr_varHorses(i, 3) 'Row number on which the horse is running
        g_arr_varReplay_HorseData(i, 11) = g_arr_varHorses(i, 11) 'Starting number
        g_arr_varReplay_HorseData(i, 15) = g_arr_varHorses(i, 15) 'Starting gate
        g_arr_varReplay_HorseData(i, 23) = g_arr_varHorses(i, 23) 'Picture of the winner
        g_arr_varReplay_HorseData(i, 24) = g_arr_varHorses(i, 24) 'For SPECIAL purposes
        g_arr_varReplay_HorseData(i, 25) = g_arr_varHorses(i, 25) 'Tactics
    Next i

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "StoreRaceReplay1()")
    Call basAuxiliary.CodeCrash
End Sub

'Store the data of the last race for replaying (part 2, after the race)
Private Sub StoreRaceReplay2(loops As Long)

    On Error GoTo ERRORHANDLING 'In case an error occurs

    'Resize the race speed log arrays to the perfect length and store horse data
    Dim arrTempSpeedLogLength() As Double
    For i = 1 To g_arr_varReplay_RaceData(15, 1) 'NUMBER_ENROLLED
    
        arrTempSpeedLogLength = g_arr_varHorses(i, 14)
        ReDim Preserve arrTempSpeedLogLength(1 To loops)

        g_arr_varReplay_HorseData(i, 14) = arrTempSpeedLogLength 'Race speed log (array with each step)
        g_arr_varReplay_HorseData(i, 20) = g_arr_varHorses(i, 20) 'Status after the race

    Next i

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "StoreRaceReplay2()")
    Call basAuxiliary.CodeCrash
    
End Sub

Public Sub ChangeColourMode(cmode As String)
    If objRace.STARTED Then 'Leave the current race and paint the start screen?
        Call ShowMessagePopup(g_c_tool, GetText(g_arr_Text, "WARN003"), _
           enumButton.CancelOK, vbModal)
        'Evaluate the return value
        If g_enumButton = enumButton.Cancel Then Exit Sub
    End If
    
    'Delete speed chart
    Dim sh As Shape
    For Each sh In g_wksRace.Shapes
       If sh.name = "SpeedChart" Then sh.Delete
    Next
    
    'http://www.vb-helper.com/howto_invert_color.html
    g_strColourMode = cmode '"STANDARD" "POPART" "LSD" "SMARTIES" "DARKMODE" "TV1960" "24H"

    'Override settings dependent on the colour mode
    Call GetColours_colMode
    
    If g_strPlayMode = "RS" Then
        Call RS_StartScreen 'Show the RS edition start screen
        Call RS_InactivateCommandButtons 'Deactivate some buttons
        objRace.STARTED = False
    Else
        Call TitleScreen
    End If
End Sub

'Calculate the ranking when one or more horses pass the finish line
Private Sub CalculateRanking()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    ReDim m_arr_varResultsCalc(1 To m_intHorsesFinishing, 0 To 6)
    
    Dim blnAssign As Boolean 'Assign the placement if true
    Dim intAssigned As Integer 'Number of horses which are placed
    Dim m As Integer 'Counter for entries in the calculation array

    m_intFinishLoop = m_intFinishLoop + 1 'New round of calculation
    
    'Reset variables
    blnAssign = False
    intAssigned = 0
    m = 1

    'Write the data of each horse that has finished into the calculation array
    For i = 1 To UBound(g_arr_varHorses)
        If g_arr_varHorses(i, 20) = "CALCULATION" Then
            g_arr_varHorses(i, 20) = "FINISHED" 'Set the horse�s new status
            m_arr_varResultsCalc(m, 1) = g_arr_varHorses(i, 11) 'Horse number
            m_arr_varResultsCalc(m, 2) = g_arr_varHorses(i, 1) 'Name of the horse
            m_arr_varResultsCalc(m, 3) = Round(g_arr_varHorses(i, 9) / 100, 0) 'Position (accuracy 10cm)
            m_arr_varResultsCalc(m, 4) = g_arr_varHorses(i, 2) 'Horse colour
            m_arr_varResultsCalc(m, 5) = g_arr_varHorses(i, 23) 'Photo
            m_arr_varResultsCalc(m, 6) = g_arr_varHorses(i, 24) 'For special purposes
            m = m + 1 'Next entry
        End If
    Next i
    
    'Assign placement to the finisher
    Do Until intAssigned >= UBound(m_arr_varResultsCalc) 'Until all finishers have got a placement
        For i = 1 To UBound(m_arr_varResultsCalc) 'Outer loop through the finishers
            If m_arr_varResultsCalc(i, 0) <> "PLACED" Then 'Find horses with no placement assigned yet
                For j = i To UBound(m_arr_varResultsCalc) 'Inner loop through the finishers
                    If m_arr_varResultsCalc(j, 0) <> "PLACED" Then 'Find horses with no placement assigned yet
                        If m_arr_varResultsCalc(i, 3) >= m_arr_varResultsCalc(j, 3) Then
                            'If the position is greater than or equal to the compared horse
                            blnAssign = True 'Ready to assign
                        Else
                            'If the compared horse is ahead
                            blnAssign = False
                            Exit For
                        End If
                    End If
                Next j
                
                'Do the assignment
                If blnAssign = True Then 'If ready to assign
                    
                    'Write the horse data into the array with the results
                    m_arr_varResults(m_intPlace, 0) = m_intFinishLoop 'Calculation loop in which the horse has finished
                    m_arr_varResults(m_intPlace, 2) = m_arr_varResultsCalc(i, 1) 'Horse number
                    m_arr_varResults(m_intPlace, 3) = m_arr_varResultsCalc(i, 2) 'Name of the horse
                    m_arr_varResults(m_intPlace, 4) = m_arr_varResultsCalc(i, 4) 'Horse colour
                    m_arr_varResults(m_intPlace, 5) = m_arr_varResultsCalc(i, 3) 'Position (accuracy 10cm)
                    m_arr_varResults(m_intPlace, 6) = m_arr_varResultsCalc(i, 5) 'Photo
                    m_arr_varResults(m_intPlace, 7) = m_arr_varResultsCalc(i, 6) 'For special purposes
                    
                    'Calculate the rank for this horse
                    If m_arr_varResults(m_intPlace, 0) = m_arr_varResults(m_intPlace - 1, 0) And _
                        m_arr_varResults(m_intPlace, 5) = m_arr_varResults(m_intPlace - 1, 5) Then
                            'If the position is exact the same as of the horse before
                            'assign the same rank
                            m_arr_varResults(m_intPlace, 1) = m_arr_varResults(m_intPlace - 1, 1)
                    Else
                        m_arr_varResults(m_intPlace, 1) = m_intPlace 'Assign rank
                    End If
                    
                    m_arr_varResultsCalc(i, 0) = "PLACED"
                    intAssigned = intAssigned + 1 'Increment the number of horses which are placed
                    m_intPlace = m_intPlace + 1 'Increment the placement for the next horse
                    blnAssign = False 'Reset the variable
                    Exit For 'Leave the outer loop
                End If
            End If
        Next i
    Loop
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "CalculateRanking()")
    Call basAuxiliary.CodeCrash
End Sub

'Rank the horses that did not finish
Private Sub RankNotFinished()
    For i = 1 To UBound(g_arr_varHorses)
        If g_arr_varHorses(i, 20) = "REFUSED" Or g_arr_varHorses(i, 20) = "KIDNAPPED" Then 'Find the horses that have not finished
            For j = 1 To UBound(m_arr_varResults) 'Find the next free line in the ranking list
                If m_arr_varResults(j, 1) = "" Then
                    m_arr_varResults(j, 1) = "-" 'No ranking
                    m_arr_varResults(j, 2) = g_arr_varHorses(i, 11) 'Horse number
                    m_arr_varResults(j, 3) = g_arr_varHorses(i, 1) 'Name of the horse
                    m_arr_varResults(j, 4) = g_arr_varHorses(i, 2) 'Horse colour
                    Exit For
                End If
            Next j
        End If
    Next i
End Sub

'Write the placements to the main array
Private Sub TransmitPlacements()
    For i = 1 To UBound(m_arr_varResults)
        For j = 1 To UBound(g_arr_varHorses)
            If m_arr_varResults(i, 2) = g_arr_varHorses(j, 11) Then 'Compare starting numbers
                g_arr_varHorses(j, 12) = m_arr_varResults(i, 1) 'Transfer the placement
            End If
        Next j
    Next i
End Sub

'Formattings for the photo of the finish and the ranking list
Private Sub FormatPhotoAndRanking()

    'Texts in case of a photo finish
    Cells(2 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)) _
        .Font.name = "Arial Black" '"PHOTO FINISH!"
    With Cells(4 + objBasicData.TOP_ROWS, _
            objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)) '"Photo..."
        .Font.size = 14
        .Font.Bold = True
    End With

    'Column width
    Range(Columns(objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), _
        Columns(objRace.METRES + objBasicData.LEFT_COLS + 175 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR))) _
        .ColumnWidth = objBasicData.RANK_WIDTH / 10 'Columns for the photo of the finish
    Columns(objRace.METRES + objBasicData.LEFT_COLS + 176 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR)) _
        .ColumnWidth = objBasicData.RANK_WIDTH   'Column behind the photo of the finish

End Sub

'Create a photo of the finish when the first horse
'crosses the finish line
Private Sub CreateFinishPhoto(replay As Boolean)
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Call basAuxiliary.ActivateRaceSheet
    Call FormatPhotoAndRanking
    
    'Write the data for the photo into an array
    For j = 1 To UBound(m_arr_varPhotofinish)
        m_arr_varPhotofinish(j, 0) = g_arr_varHorses(j, 3) 'Row number on which the horse is running
        m_arr_varPhotofinish(j, 1) = Round(g_arr_varHorses(j, 9) / 100, 0) 'Position (accuracy 10cm)
        m_arr_varPhotofinish(j, 2) = g_arr_varHorses(j, 11) 'Horse number
        m_arr_varPhotofinish(j, 4) = g_arr_varHorses(j, 24) 'For special purposes
    Next j
    
    'Flash in case of a tight finish
    If m_blnPhotofinish Then

        If Not replay Then
            'Announcement "PHOTO FINISH!"
            If objOption.SPEECH Then Call SpeechOut(GetText(g_arr_Text, "RACE025")) 'Voice output if selected
            With Cells(2 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))
                .Value = GetText(g_arr_Text, "RACE025")
                .Font.Color = objOption.COL_TEXT
            End With
        End If
            
        'Alternate the track colour rapidly between black and white behind the finish line
        For k = 1 To 8 'Run the loop 8 times
            With Range(Cells(5 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), _
                Cells(objRace.NUMBER_ENROLLED * 2 + 7 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)))
                    .Interior.ColorIndex = 1 'Black
                    .Interior.ColorIndex = 0 'White
            End With
        Next k
        'Reset the track to its original colour
        With Range(Cells(5 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), _
            Cells(objRace.NUMBER_ENROLLED * 2 + 7 + objBasicData.TOP_ROWS, objRace.METRES + objBasicData.LEFT_COLS + 14 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)))
                .Interior.Color = objRace.TRACK_COLOUR
                If objRace.SPECIAL = "PARTICULATES" Then .Interior.Pattern = objOption.PARTICULATES_PATTERN
        End With
    End If

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "CreateFinishPhoto()")
    Call basAuxiliary.CodeCrash
End Sub

'Draw the photo of the finish
Private Sub DrawFinishPhoto()
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Call basAuxiliary.ActivateRaceSheet
    
    'Prepare a variable of type Long, otherwise an overflow occurs in races with long distances
    'when multiplying the track length for calculating the exact position
    Dim lngFin As Long
    lngFin = objRace.METRES 'Copy the track length into the variable
        
    'Prepare variables for colours
    Dim colourTrack As Long 'Background of the photo
    Dim colourLines As Long 'Lines on the photo
    Dim colourNames As Long 'Horse names
    Dim colourScale As Long 'Stripe with the metre scale
    
    If objOption.PHOTO_BW Then 'If the photo is to be displayed in black-and-white
        colourTrack = 1 'Black track
        colourLines = 2 'White lines
        colourNames = vbWhite
        colourScale = vbWhite
    Else
        colourTrack = objRace.TRACK_COLOUR 'Original track colour
        colourLines = 1 'Black lines
        colourNames = objOption.COL_TEXT
        colourScale = objRace.TRACK_COLOUR
    End If
    
    Application.ScreenUpdating = False 'Deactivate screen updating
    
    'Clear the range on which the photo is to be shown
    Range(Cells(5 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 8 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 175 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR))) _
            .Clear
    
    'Draw a frame around the photo
    Range(Cells(5 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 8 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 175 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR))) _
            .BorderAround Color:=objOption.COL_TEXT, Weight:=xlMedium
            
    'Draw a horizontal line for the section with the scale markers
    With Range(Cells(objRace.NUMBER_ENROLLED * 2 + 7 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 7 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 175 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR))) _
            .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = colourLines
                .Weight = xlThin
    End With
    
    'Write the race data as a caption
    With Cells(5 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))
            .Font.name = "Arial"
            .Font.Bold = True
            .Font.Color = colourNames
            .Value = objRace.RACE_NAME & " " & objRace.RACE_YEAR _
                & " - " & objRace.TRACK_NAME & ", " & objRace.TRACK_LOCATION _
                & " - " & objRace.METRES & " " & GetText(g_arr_Text, "RACE009")
    End With

    'Draw the race track and the finish line
    Range(Cells(5 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 7 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 175 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR))) _
            .Interior.Color = colourTrack 'Background
    Range(Cells(objRace.NUMBER_ENROLLED * 2 + 8 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 8 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 175 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR))) _
            .Interior.Color = colourScale 'Stripe for the scale markers
    Range(Cells(5 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 155 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 7 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 155 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))) _
        .Interior.ColorIndex = colourLines 'Finish line
        
    'Draw track markers ("XXX metres")
    For i = 1 To (15 + 2 * objOption.SPEED_FACTOR)
        Cells(objRace.NUMBER_ENROLLED * 2 + 7 + objBasicData.TOP_ROWS, _
            i * 10 + objRace.METRES + objBasicData.LEFT_COLS + 15 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)) _
            .Interior.ColorIndex = colourLines 'Scale (vertical bars)
        With Cells(objRace.NUMBER_ENROLLED * 2 + 8 + objBasicData.TOP_ROWS, _
            i * 10 + objRace.METRES + objBasicData.LEFT_COLS + 13 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))
            .Value = (objRace.METRES - 14 + i) & GetText(g_arr_Text, "RACE008")
            .Font.Color = vbBlack 'Scale (metres)
        End With
    Next i

    'Prepare variables
    Dim photoLeftMargin As Integer 'Column of the left edge of the photo
    Dim racetrackMinimumColumn As Long 'Minimum horse head position to appear on the photo
    Dim currentDrawColumn As Integer 'Column for drawing the current segment

    photoLeftMargin = objRace.METRES + objBasicData.LEFT_COLS + 16 _
        + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR) 'Get the left column of the photo
    racetrackMinimumColumn = objRace.METRES * 10 - 139 'Get the minimum position to appear
                
    'Paint the horses
    For i = 1 To UBound(m_arr_varPhotofinish)
        If m_arr_varPhotofinish(i, 1) >= racetrackMinimumColumn Then 'Only if the horse appears in the photo (at least partially)
            currentDrawColumn = photoLeftMargin + m_arr_varPhotofinish(i, 1) - racetrackMinimumColumn 'Column of the horse�s head
          
            'Find the next horse to paint (sort by horse numbers ascending)
            For j = 1 To UBound(m_arr_varResults)
                If m_arr_varResults(j, 2) = m_arr_varPhotofinish(i, 2) Then
                    Exit For
                End If
            Next j
            
            'Draw a vertical line at the position of the horse�s head
            '(only for those which have crossed the finish line)
            If m_arr_varPhotofinish(i, 1) >= (objRace.METRES * 10) Then
                With Range(Cells(5 + objBasicData.TOP_ROWS, currentDrawColumn), _
                    Cells(objRace.NUMBER_ENROLLED * 2 + 8 + objBasicData.TOP_ROWS, currentDrawColumn)) _
                        .Borders(xlEdgeRight)
                            .LineStyle = xlContinuous
                            .ColorIndex = 0
                            .Weight = xlHairline 'Very thin line
                End With
            End If
            
            'Prepare variables for the horse segments
            Dim horseSegment As Integer 'Number of the current segment (1=tail, 8=head)
            Dim horseSegmentColour As Long 'Colour of the current segment
            Dim segmentLength As Integer 'Length of the current segment
            
            For horseSegment = 8 To 1 Step -1 'Loop through the horse segments starting from the head
                If IsArray(m_arr_varResults(j, 4)) Then
                    'Multicoloured horse: Get the colour from the array
                    horseSegmentColour = m_arr_varResults(j, 4)(horseSegment - 1)
                Else
                    'Monochrome horse (no array submitted)
                    horseSegmentColour = m_arr_varResults(j, 4)
                End If
                
                'Convert to grey in case of a photo in black-and-white
                If objOption.PHOTO_BW Then horseSegmentColour = GreyToLong(CInt(RGBtoGrey(CLng(horseSegmentColour))))
                
                If currentDrawColumn >= photoLeftMargin Then 'Only if still visible on the photo
                    If currentDrawColumn - photoLeftMargin >= 10 Then
                        segmentLength = 10 'Complete segment visible
                    Else
                        segmentLength = currentDrawColumn - photoLeftMargin + 1 'Segment only partially visible
                    End If
                    
                    'Paint the segment
                    Range(Cells(m_arr_varPhotofinish(i, 0), currentDrawColumn), _
                          Cells(m_arr_varPhotofinish(i, 0), currentDrawColumn - segmentLength + 1)) _
                          .Interior.Color = horseSegmentColour
                    
                    currentDrawColumn = currentDrawColumn - segmentLength 'Adjust the column for the next segment
                End If
            Next horseSegment 'Next segment
        End If
        
        'Write the horse names in the photo (if selected in the race options)
        If objOption.NAMES_PHOTO = True And g_arr_varHorses(i, 0) = "START" Then
            With Cells(g_arr_varHorses(i, 3), _
                    objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))
                .Font.name = "Arial"
                .Font.size = m_intFontSize
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .Font.Color = colourNames
                .Value = g_arr_varHorses(i, 1) 'Horse name
            End With
        End If
    Next i
    
    If objOption.AUTOFIT Then Call AutoZoom("FinishPhoto")
    Application.ScreenUpdating = True 'Activate screen updating
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "DrawFinishPhoto()")
    Call basAuxiliary.CodeCrash
End Sub

'Race is finished - show info pop-up
Private Sub RaceFinished()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Dim messagetext As String
    messagetext = GetText(g_arr_Text, "RACE028") & vbNewLine & GetText(g_arr_Text, "RACE029")
    
    'Pop-up
    If objOption.SPEECH Then Call SpeechOut(messagetext) 'Voice output if selected
        
    Call ShowMessagePopup(objRace.RACE_NAME & " " & objRace.RACE_YEAR, _
        messagetext, enumButton.OK, vbModal)
        
    'Unfreeze the window pane
    If objOption.NAMES_LEFT Or objOption.COLOURS_LEFT Or objOption.HIGHLIGHT_FAV _
        Or objOption.MOMENTUM_BARS Or objOption.MOMENTUM_ICONS Or objOption.TACTICS_REVEAL_TAC _
        Or objOption.TACTICS_REVEAL_CURR Or (objOption.FOCUSED_RUN <> enumCamera.standard And objOption.HIGHLIGHT_FOC) _
        Or (objOption.RACE_INFO And objOption.RACE_INFO_WKS) Then Call basAuxiliary.Freeze(0, 0, False)
        
    'Scrollen zu Ergebnistafel
    Call basAuxiliary.Scroll(objRace.METRES + objBasicData.LEFT_COLS + 15 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR), objBasicData.TOP_ROWS + (objRace.NUMBER_ENROLLED * 2 + 9))

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RaceFinished()")
    Call basAuxiliary.CodeCrash
End Sub

'Draw the ranking list
Private Sub DrawRankingList(afterRace As Boolean, test As Boolean, replay As Boolean)
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Call basAuxiliary.ActivateRaceSheet
    If Not replay Then Call basAuxiliary.Scroll(objRace.METRES + objBasicData.LEFT_COLS + 15 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR), objBasicData.TOP_ROWS + (objRace.NUMBER_ENROLLED * 2 + 9))

    'Scoreboard
    With Range(Cells(objRace.NUMBER_ENROLLED * 2 + 20 + objBasicData.TOP_ROWS, _
            objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), _
            Cells(objRace.NUMBER_ENROLLED * 2 + 20 + objRace.NUMBER_STARTING + 1 + objBasicData.TOP_ROWS, _
            objRace.METRES + objBasicData.LEFT_COLS + 175 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)))
        .Clear 'clear all cell values and formattings
        .BorderAround Color:=objOption.COL_TEXT, Weight:=xlMedium  'Border
        .Interior.Color = objOption.COL_RANKINGS 'Background
        With .Font
            .name = "Courier New"
            .size = 12
            .Color = objOption.COL_TEXT
        End With
       .NumberFormat = "@" 'Force text format
    End With
    With Cells(objRace.NUMBER_ENROLLED * 2 + 20 + objBasicData.TOP_ROWS, _
                objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))
        .Font.size = 14 'Formattings for the headline
        .Font.Bold = True
        .IndentLevel = 1
    End With
    
    'Headline
    Cells(objRace.NUMBER_ENROLLED * 2 + 20 + objBasicData.TOP_ROWS, _
            objRace.METRES + objBasicData.LEFT_COLS + 16 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)) _
            .Value = objRace.RACE_NAME & " " & objRace.RACE_YEAR & " - " & objRace.TRACK_LOCATION
            
    If Not g_skipDelay And objOption.RANKING_DELAY And afterRace Then _
        Application.Wait (Now + TimeValue("0:00:04")) 'Delay
        
    'Results
    Dim intPositionName As Integer 'Position of the horse names
    intPositionName = 0
    For i = UBound(m_arr_varResults) To 1 Step -1 'Loop through through the results bottom up
        If objOption.RANKING_COL Then 'Show colours if selected in the race options
            intPositionName = 12
            Call PaintHorse(objRace.NUMBER_ENROLLED * 2 + 20 + i + objBasicData.TOP_ROWS, _
                objRace.METRES + objBasicData.LEFT_COLS + 19 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR), _
                m_arr_varResults(i, 4)) 'Paint the horse
            Range(Cells(objRace.NUMBER_ENROLLED * 2 + 20 + i + objBasicData.TOP_ROWS, _
                objRace.METRES + objBasicData.LEFT_COLS + 19 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)), _
                Cells(objRace.NUMBER_ENROLLED * 2 + 20 + i + objBasicData.TOP_ROWS, _
                objRace.METRES + objBasicData.LEFT_COLS + 26 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))) _
                .BorderAround ColorIndex:=0, Weight:=xlThin 'Draw a frame around the horse
        End If
        Cells(objRace.NUMBER_ENROLLED * 2 + 20 + i + objBasicData.TOP_ROWS, _
            objRace.METRES + objBasicData.LEFT_COLS + 22 + intPositionName + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)) _
            .Value = m_arr_varResults(i, 1) & "."  'Place
        Cells(objRace.NUMBER_ENROLLED * 2 + 20 + i + objBasicData.TOP_ROWS, _
            objRace.METRES + objBasicData.LEFT_COLS + 29 + intPositionName + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR)) _
            .Value = m_arr_varResults(i, 3) & " (#" & m_arr_varResults(i, 2) & ")" 'Horse name and number
        
    'Check if this is a Matjes Grand Prix (GMP) race
    If GMP_mode = "QUALIFICATION" Then 'GMP qualification race
        If objRace.NUMBER_STARTING > 8 Then
            GMP_qualified = 3 'Three horses qualifiy for semifinals
        Else
            GMP_qualified = 2 'Two horses qualifiy for semifinals
        End If
        If m_arr_varResults(i, 1) <= GMP_qualified Then
            Call GMP_Ranking(intPositionName, True, " >>> " & GetText(g_arr_Text, "GMP01"))
        Else
            Call GMP_Ranking(intPositionName, False)
        End If
    ElseIf GMP_mode = "SEMIFINAL" Then 'GMP semifinal race
        GMP_qualified = 6 'Six horses qualifiy for the final race
        If m_arr_varResults(i, 1) <= GMP_qualified Then
            Call GMP_Ranking(intPositionName, True, " >>> " & GetText(g_arr_Text, "GMP02"))
        Else
            Call GMP_Ranking(intPositionName, False)
        End If
    Else 'No GMP race, GMP training or final
        Call GMP_Ranking(intPositionName, False)
    End If
    
    'Check if this is a Corona Cup race
    If GMP_mode = "CC21" Then 'Corona Cup qualification race
    
        CC_qualifed = 4 'Three horses qualifiy for the final
        
        If m_arr_varResults(i, 1) <= CC_qualifed Then
            Call CC_Ranking(intPositionName, True, " >>> " & GetText(g_arr_Text, "CC01"))
            If g_strPlayMode = "RS" Then Call CC_SetupFinal
        Else
            Call CC_Ranking(intPositionName, False)
        End If
        
    Else 'Corona Cup final
        Call CC_Ranking(intPositionName, False)
    End If
        
        If Not g_skipDelay And objOption.RANKING_DELAY And afterRace Then _
            Application.Wait (Now + TimeValue("0:00:01")) 'Delay
    Next i
                
    'Show a pop-up in case of a dead heat (more than one winner)
    If m_blnDeadHeat And Not test And Not replay Then ShowDeadHeat
    'Alternative variants:
'    If replay = False And test = False And m_blnDeadHeat = True Then Call ShowDeadHeat
'    If replay = False And test = False And m_blnDeadHeat = True Then ShowDeadHeat
'    If replay = False And test = False And m_blnDeadHeat Then Call ShowDeadHeat

    If objOption.AUTOFIT And Not replay Then Call AutoZoom("RankingList")
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "DrawRankingList()")
    Call basAuxiliary.CodeCrash
End Sub

'Extend the score board in case of a Matjes Grand Prix (GMP) race
Sub GMP_Ranking(intPositionName As Integer, q1 As Boolean, Optional q2 As String)
            With Cells(objRace.NUMBER_ENROLLED * 2 + 20 + i + objBasicData.TOP_ROWS, _
                objRace.METRES + objBasicData.LEFT_COLS + 29 + intPositionName + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))
                .Value = m_arr_varResults(i, 3) & " (#" & m_arr_varResults(i, 2) & ")" & q2
                If q1 = True Then
                .Font.Color = RGB(0, 135, 60)
                .Font.Bold = True
                End If
            End With
End Sub

'Extend the score board in case of a Corona Cup race
Sub CC_Ranking(intPositionName As Integer, q1 As Boolean, Optional q2 As String)
            With Cells(objRace.NUMBER_ENROLLED * 2 + 20 + i + objBasicData.TOP_ROWS, _
                objRace.METRES + objBasicData.LEFT_COLS + 29 + intPositionName + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR))
                .Value = m_arr_varResults(i, 3) & " (#" & m_arr_varResults(i, 2) & ")" & q2
                If q1 = True Then
                .Font.Color = RGB(0, 135, 60)
                .Font.Bold = True
                End If
            End With
End Sub

'Copy the qualified horses to the race spreadsheet with the final
Sub CC_SetupFinal()

    Dim wksFin As Worksheet
    Dim nextLine As Integer, nextColumn As Integer
    Dim originalLine As Integer
    
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Set wksFin = ThisWorkbook.Worksheets("race_CORONA21F") 'Get the Corona Cup Final race sheet
    nextLine = wksFin.Cells(rows.count, 5).End(xlUp).row + 1 'Find the next free line
    
    Do
        originalLine = originalLine + 1
        If g_wksRaceData.Cells(originalLine, 7).Value = m_arr_varResults(i, 3) Then Exit Do 'Line found
    Loop

    wksFin.Cells(nextLine, 5).Value = nextLine - 1 'Assign the next free start number
    
    'Copy the horse data
    For nextColumn = 6 To 20
        wksFin.Cells(nextLine, nextColumn).Value = g_wksRaceData.Cells(originalLine, nextColumn).Value
    Next nextColumn
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "CC_SetupFinal()")
End Sub

'Speed factor for a single step (recalculated for each horse in every loop)
Function SpeedLoop() As Double
    Randomize 'Initialize the random number generator
    'Get a random number within a defined range
    SpeedLoop = Int((objSpeed.SPEED_LOOP_HIGH - objSpeed.SPEED_LOOP_LOW + 1) * Rnd + objSpeed.SPEED_LOOP_LOW)
End Function

'Display the winner
Private Sub DrawWinnerPhoto(Optional zoom As Boolean)
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Call basAuxiliary.ActivateRaceSheet
    Application.ScreenUpdating = False 'Deactivate screen updating
    
    'Formattings for the name of the winner
    With Range(Cells(objRace.NUMBER_ENROLLED * 2 + 20 + objBasicData.TOP_ROWS, _
            objRace.METRES + objBasicData.LEFT_COLS + 177 + (2 * 10 * objOption.SPEED_FACTOR)), _
            Cells(objRace.NUMBER_ENROLLED * 2 + 21 + objBasicData.TOP_ROWS, _
            objRace.METRES + objBasicData.LEFT_COLS + 179 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR)))
        With .Font
            .Color = objOption.COL_TEXT
            .size = 14
            .Bold = True
        End With
    End With
    
    'Reset the number of winners
    objStat.WINNERS = 0
    
    For i = 1 To UBound(m_arr_varResults)
        If m_arr_varResults(i, 1) = 1 Then

            'Column width for the photo
            Range(Columns(objRace.METRES + objBasicData.LEFT_COLS + 177 + ((i - 1) * 21) + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR)), _
            Columns(objRace.METRES + objBasicData.LEFT_COLS + 197 + ((i - 1) * 21) + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR))) _
            .ColumnWidth = 2

            'Draw the photo (20 columns x 18 rows)
            Call PaintPicture(g_wksPIC, g_wksRace, m_arr_varResults(i, 6), 20, 18, _
                objBasicData.TOP_ROWS + objRace.NUMBER_ENROLLED * 2 + 23, _
                objBasicData.LEFT_COLS + objRace.METRES + 177 + ((i - 1) * 21) + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR))

            'Draw a frame around the photo
            Range(Cells(objRace.NUMBER_ENROLLED * 2 + 23 + objBasicData.TOP_ROWS, _
                objRace.METRES + objBasicData.LEFT_COLS + 177 + ((i - 1) * 21) + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR)), _
                Cells(objRace.NUMBER_ENROLLED * 2 + 40 + objBasicData.TOP_ROWS, _
                objRace.METRES + objBasicData.LEFT_COLS + 196 + ((i - 1) * 21) + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR))) _
                .BorderAround Color:=objOption.COL_TEXT, Weight:=xlMedium 'Draw a frame around the horse
        
            'Count the number of winners
            objStat.WINNERS = objStat.WINNERS + 1
        
        End If
    Next i
    
    Cells(objRace.NUMBER_ENROLLED * 2 + 20 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 177 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR)) _
        .Value = GetText(g_arr_Text, "RACE031") '"Winner of the race:"
    Cells(objRace.NUMBER_ENROLLED * 2 + 21 + objBasicData.TOP_ROWS, _
        objRace.METRES + objBasicData.LEFT_COLS + 179 + objBasicData.AFTER_FIN_COLS + (2 * 10 * objOption.SPEED_FACTOR)) _
        .Value = objStat.WINNER_NAME 'Name of the winner

    If objOption.AUTOFIT And zoom Then Call AutoZoom("WinnerPhoto")
    
    Application.ScreenUpdating = True 'Activate screen updating
    
    'Voice output if selected in the race options
    If objOption.SPEECH Then Call SpeechOut(GetText(g_arr_Text, "RACE031"))
    If objOption.SPEECH Then Call SpeechOut(objStat.WINNER_NAME)
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "DrawWinnerPhoto()")
    Call basAuxiliary.CodeCrash
End Sub

'Check for a dead heat (i.e. more than one winner)
Private Sub CheckDeadHeat()
    'Reset variables
    m_blnDeadHeat = False
    objStat.WINNER_NAME = ""
    
    'Loop through the array with the race results
    For i = 1 To UBound(m_arr_varResults)
        If m_arr_varResults(i, 1) = 1 Then 'Find horses that rank 1st
            If i > 1 Then
                'In case of more than one winner
                objStat.WINNER_NAME = objStat.WINNER_NAME & " " & GetText(g_arr_Text, "RACE017") & " " 'add " and "
                m_blnDeadHeat = True 'Flag for a dead heat
            End If
            'Compile the string with the name(s) of the winner(s)
            objStat.WINNER_NAME = objStat.WINNER_NAME & UCase(m_arr_varResults(i, 3))
        End If
    Next i
    
End Sub

'Show a pop-up in case of a dead heat
Private Sub ShowDeadHeat()
    Call ShowInfoPopup(objRace.RACE_NAME & " " & objRace.RACE_YEAR, _
        UCase(GetText(g_arr_Text, "RACE033")) & "!", _
        False, vbModal, 22)
End Sub

'Analyse bettings
Private Sub AnalyseBettings()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    'Variable for counting the number of winners in a dead heat
    Dim intDeadHeatWinners As Integer
    
    'Variables for a single bet slip
    Dim id As String
    Dim nm As String
    Dim ty As Long
    Dim ty_txt As String
    Dim St As Double
    Dim od As Double
    Dim bt() As Integer
    Dim payout As Boolean
    Dim payCash As String 'Pay-out value
    Dim payColor As Long 'Colour of the pay-out label
        
    'Variable for text label alignment
    Dim align As Integer
    
    'Variables for statistical purposes
    Dim totalStake As Double
    Dim totalPayout As Double
    
    If g_colBetSlips.count > 0 Then 'Only if bets have been placed for this race

        'Check how many horses have won in case of a dead heat
        If m_blnDeadHeat Then
            For i = 1 To UBound(m_arr_varResults)
                If m_arr_varResults(i, 1) = 1 Then intDeadHeatWinners = intDeadHeatWinners + 1
            Next
        
            #If Debugging Then
                Debug.Print vbNewLine & "Dead Heat - " & intDeadHeatWinners & " winners"
            #End If

        End If
        
        'Calculate the pay-out for bet type 2 sur 4
        Dim payout2sur4 As Double
        For i = 1 To 4 'Loop through the top 4
            For j = 1 To UBound(g_arr_varHorses)
                If m_arr_varResults(i, 2) = g_arr_varHorses(j, 11) Then 'Find the horse
                    payout2sur4 = payout2sur4 + g_arr_varHorses(j, 17) 'Sum up the odds
                End If
            Next
        Next
        payout2sur4 = Round(((payout2sur4 / 4) / objRace.NUMBER_STARTING * 18 / 7), 0)

        'Create a label with the headline for the racing results
        Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
        With g_objLabel
            With .Font
                .name = "Tahoma"
                .size = 12
                .Bold = True
            End With
            .left = 15
            .top = 10
            .width = 300
            .TextAlign = fmTextAlignLeft
            .caption = GetText(g_arr_Text, "BET039") '"Official racing result"
            If m_blnDeadHeat Then .caption = .caption & " // " & UCase(GetText(g_arr_Text, "RACE033") & "!") '"DEAD HEAT"
        End With
    
        'Create a label with the horses that finished on place 1-4
        Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
        With g_objLabel
            .Font.name = "Tahoma"
            .Font.size = 10
            .left = 40
            .top = 30
            .width = 300
            .Height = 50
            .TextAlign = fmTextAlignLeft
            For i = 1 To 4 'Compile the label text with the horses on place 1-4
                .caption = .caption & GetText(g_arr_Text, "BET040") & " " & m_arr_varResults(i, 1) & ": " _
                                & m_arr_varResults(i, 3) & " (#" & m_arr_varResults(i, 2) & ")" & vbNewLine
            Next i
        End With
    
        'Create a label with the headline for the placed bets
        Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
        With g_objLabel
            With .Font
                .name = "Tahoma"
                .size = 12
                .Bold = True
            End With
            .left = 15
            .top = 90
            .width = 300
            .TextAlign = fmTextAlignLeft
            .caption = GetText(g_arr_Text, "BET042") '"Placed bets"
        End With
        
        align = 1
             
        'Loop through the array with the placed bets
        For i = 1 To g_colBetSlips.count
            payout = False 'Set the initial value for pay-out
            'Get the data of a bet
            id = g_colBetSlips(i).BET_ID
            nm = g_colBetSlips(i).GAMBLERNAME
            ty = g_colBetSlips(i).BET_TYPE
            ty_txt = g_colBetSlips(i).BET_TYPE_TEXT
            St = g_colBetSlips(i).STAKE
            od = g_colBetSlips(i).ODDS * 10
            bt() = g_colBetSlips(i).BET 'Prediction of the ranking (place 1-4, dependent on the type of the bet)
            
            'Analyse the bet slips by comparing the prediction with the actual rank
            Dim x As Integer, y As Integer, z As Integer
            Select Case ty

                Case enumBetType.win 'Type of bet: Win
                    For x = 1 To UBound(m_arr_varResults)
                        If bt(1) = m_arr_varResults(x, 2) _
                            And m_arr_varResults(x, 1) = 1 Then
                                payout = True
                                Exit For
                        End If
                    Next

                Case enumBetType.show
                    If objRace.NUMBER_STARTING >= 12 Then '12 or more horses running
                        y = 4
                    ElseIf objRace.NUMBER_STARTING >= 8 Then '8-11 horses
                        y = 3
                    Else 'up to 7 horses
                        y = 2
                    End If
                    For x = 1 To UBound(m_arr_varResults)
                        If bt(1) = m_arr_varResults(x, 2) _
                            And m_arr_varResults(x, 1) <= y Then
                            payout = True
                            Exit For
                        End If
                    Next

                Case enumBetType.x2sur4
                    z = 0
                    For x = 1 To UBound(m_arr_varResults)
                        For y = 1 To 2
                            If bt(y) = m_arr_varResults(x, 2) _
                                And m_arr_varResults(x, 1) <= 4 Then
                                    z = z + 1
                                    Exit For
                            End If
                        Next
                    Next
                    od = payout2sur4 'Assign the pay-out
                    If z = 2 Then payout = True

                Case enumBetType.exacta
                    z = 0
                    For x = 1 To UBound(m_arr_varResults)
                        For y = 1 To 2
                            If bt(y) = m_arr_varResults(x, 2) _
                                And m_arr_varResults(x, 1) <= y Then
                                    z = z + 1
                                    Exit For
                            End If
                        Next
                    Next
                    If z = 2 Then payout = True

                Case enumBetType.trifecta
                    z = 0
                    For x = 1 To UBound(m_arr_varResults)
                        For y = 1 To 3
                            If bt(y) = m_arr_varResults(x, 2) _
                                And m_arr_varResults(x, 1) <= y Then
                                    z = z + 1
                                    Exit For
                            End If
                        Next
                    Next
                    If z = 3 Then payout = True
                    
                Case enumBetType.superfecta
                    z = 0
                    For x = 1 To UBound(m_arr_varResults)
                        For y = 1 To 4
                            If bt(y) = m_arr_varResults(x, 2) _
                                And m_arr_varResults(x, 1) <= y Then
                                    z = z + 1
                                    Exit For
                            End If
                        Next
                    Next
                    If z = 4 Then payout = True

            End Select
            
            'Set the value and the colour of the pay-out label
            If payout = True And m_blnDeadHeat And ty = enumBetType.win Then
                payCash = Round(St / 10 * od / intDeadHeatWinners, 1) 'Pay-out (dead heat)
                payColor = 52377 'Green
            ElseIf payout = True Then
                payCash = St / 10 * od 'Pay-out (full)
                payColor = 52377 'Green
            Else
                payCash = 0 'No pay-out
                payColor = &H8080FF 'Red
            End If
            
            'Write the name of the gambler and the bet slip ID
            Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
            With g_objLabel
                With .Font
                    .name = "Tahoma"
                    .size = 10
                    .Bold = True
                End With
                .left = 40
                .top = 98 + align * 12
                .width = 350
                .TextAlign = fmTextAlignLeft
                .caption = nm & " (" & GetText(g_arr_Text, "BET001") & " " & GetText(g_arr_Text, "ODDS001") & " " & id & ")"
            End With
            
            align = align + 1
            
            'Write the type of the bet
            Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
            With g_objLabel
                .Font.name = "Tahoma"
                .Font.size = 10
                .left = 80
                .top = 98 + align * 12
                .width = 200
                .TextAlign = fmTextAlignLeft
                .caption = UCase(GetText(g_arr_Text, "BET007")) & ": " & ty_txt '"TYPE OF BET:"
            End With
        
            align = align + 1
            
            'Write the horse name and the predicted rank
            For j = 1 To UBound(bt)
                Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
                With g_objLabel
                    .Font.name = "Tahoma"
                    .Font.size = 10
                    .left = 80
                    .top = 98 + align * 12
                    .width = 200
                    .TextAlign = fmTextAlignLeft
                    .caption = GetHorseName(bt(j)) & " (#" & bt(j) & ")"
                End With
                
                align = align + 1
            Next j
            
            'Write the stake and currency
            Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
            With g_objLabel
                .Font.name = "Tahoma"
                .Font.size = 10
                .left = 80
                .top = 98 + align * 12
                .width = 100
                .TextAlign = fmTextAlignLeft
                .caption = WorksheetFunction.Proper(GetText(g_arr_Text, "BET036")) & ": " & Format(St, "0.00") & " " & GetText(g_arr_Text, "BET035") '"Stake: x.xx EUR"
            End With
                
            'Write the pay-out and currency
            Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
            With g_objLabel
                .Font.name = "Tahoma"
                .Font.size = 10
                .left = 200
                .top = 98 + align * 12
                .width = 150
                .TextAlign = fmTextAlignLeft
                .caption = "  " & GetText(g_arr_Text, "BET038") & ": " & Format(payCash, "#,##0.00") & " " & GetText(g_arr_Text, "BET035") 'Pay-out: xx EUR
                .BackColor = payColor
            End With
                
            'For statistical purposes: Calculate the total stakes and pay-out
            totalStake = totalStake + St
            totalPayout = totalPayout + payCash
            
            align = align + 2
            
            'For statistical purposes: Payout logging (single bet)
            If g_payoutLogging And objOption.BET_PLACED Then
                Open Environ("UserProfile") & "\Desktop\GALOPPSIM_PAYOUTLOG.csv" For Append As #1
                    Print #1, Date & ";" & "Single bet" & ";" & objRace.RACE_ID & ";" _
                        & objRace.NUMBER_STARTING & ";" & ";" & ty_txt _
                        & ";" & Format(St, "0.00") & ";" & Format(payCash, "0.00")
                Close #1
            End If
            
        Next i 'Next bet
        
        'Write the total number of placed bets for this race
        Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
        With g_objLabel
            With .Font
                .name = "Tahoma"
                .size = 10
            End With
            .left = 40
            .top = 98 + align * 12
            .width = 300
            .TextAlign = fmTextAlignLeft
            .caption = GetText(g_arr_Text, "START012") & ": " & g_colBetSlips.count '"Number of bet slips:"
        End With
            
        align = align + 1
        
        'Write the total stakes and pay-out for this race
        Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
        With g_objLabel
            With .Font
                .name = "Tahoma"
                .size = 10
            End With
            .left = 40
            .top = 98 + align * 12
            .width = 350
            .TextAlign = fmTextAlignLeft
            .caption = GetText(g_arr_Text, "BET043") & ": " & Format(totalStake, "#,##0.00") & " " & GetText(g_arr_Text, "BET035") & " / " _
                        & GetText(g_arr_Text, "BET044") & ": " & Format(totalPayout, "#,##0.00") & " " & GetText(g_arr_Text, "BET035")
        End With

        'Pop-up formattings
        With frmBettingAnalysis
            .caption = objRace.RACE_NAME & " " & objRace.RACE_YEAR & " | " & objRace.TRACK_NAME & ", " & objRace.TRACK_LOCATION _
                        & " (" & objRace.COUNTRY_NAME & ")" 'Race data
            .width = 400 'Fixed width of the pop-up

            If g_colBetSlips.count <= 5 Then
                .Height = 98 + align * 12 + 50 'Height of the pop-up depending of the number of placed bets
            Else
                .Height = 440 'Fixed height if more than 5 bets are placed
            End If
            .ScrollBars = fmScrollBarsVertical 'Provide a vertical scrollbar
            .ScrollHeight = 98 + align * 12 + 30 'Height of the vertical scrolling
            .KeepScrollBarsVisible = fmScrollBarsNone 'Show the scrollbar only if needed
            'Position of the pop-up on the screen
            .StartUpPosition = 0
            .top = ActiveWindow.top + ((ActiveWindow.Height - .Height) / 2) 'Vertically centred
            .left = ActiveWindow.left + ((ActiveWindow.width - .width) - ActiveWindow.width / 10) 'Near the right border
            .show (vbModeless)
        End With
        
        'For statistical purposes: Payout logging (summary per race)
        If g_payoutLogging And objOption.BET_PLACED Then
            Open Environ("UserProfile") & "\Desktop\GALOPPSIM_PAYOUTLOG.csv" For Append As #1
                Print #1, Date & ";" & "Race summary" & ";"; objRace.RACE_ID & ";" & objRace.NUMBER_STARTING & ";" _
                    & g_colBetSlips.count & ";" & ";" & Format(totalStake, "0.00") & ";" & Format(totalPayout, "0.00")
            Close #1
        End If
        
    End If

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "AnalyseBettings()")
    Call basAuxiliary.CodeCrash
End Sub

'Retrieve the horse name by the horse number
Private Function GetHorseName(num As Integer) As String
    Dim x As Integer
    For x = 1 To UBound(g_arr_varHorses())
        If num = g_arr_varHorses(x, 11) Then Exit For
    Next
    GetHorseName = g_arr_varHorses(x, 1)
End Function

'AI mode: Delete all GaloppSim worksheets
Private Sub AI_DeleteWorksheets()
    On Error Resume Next
        Application.DisplayAlerts = False 'Turn off application warnings
        'Delete worksheets
            g_wksRace.Delete
        Application.DisplayAlerts = True 'Turn on application warnings
    On Error GoTo 0
End Sub

'Provide information for the pop-up when starting a new race
Private Sub ShowStartPopup()

    With frmStart
        .caption = g_c_tool
        .lblS1.caption = objRace.RACE_NAME & " " & objRace.RACE_YEAR 'Race name and year
        .lblS2.caption = objRace.RACE_TYPE_TEXT & " " & GetText(g_arr_Text, "RACE007") & " " & objRace.METRES & " " & GetText(g_arr_Text, "RACE009") 'Race type and distance
        .lblS3.caption = objRace.TRACK_NAME & " " & GetText(g_arr_Text, "RACE002") & " " & objRace.TRACK_LOCATION & " (" & objRace.COUNTRY_NAME & ")" 'Race track, loaction and country
        .lblS6.caption = objRace.NUMBER_STARTING & " " & g_arr_Grammar(4) 'Number of horses starting
        If objRace.REAL_RACE = "Y" Then
            .lblS10.caption = UCase(GetText(g_arr_Text, "START009")) '"REAL RACE"
        Else
            .lblS10.caption = UCase(GetText(g_arr_Text, "START010")) '"FICTITIOUS RACE"
        End If
        .lblS8.caption = GetText(g_arr_Text, "START005") 'Caption of the betting section
        .lblBet02.Visible = False 'Show the label with the number of betting slips
        .lblFocus.caption = g_arr_Grammar(1) & " " & GetText(g_arr_Text, "START006") 'Label "Horse in focus"
        .lblRSMon.caption = GetText(g_arr_Text, "START016")
        .cmdS1.caption = GetText(g_arr_Text, "START002") 'Button "Add bet slip"
        .cmdS2.caption = GetText(g_arr_Text, "START003") 'Button "Start the race"
        With .lblS4 'Track surface
            .caption = objRace.TRACK_SURFACE_TEXT 'Surface type
            .BorderStyle = fmBorderStyleSingle 'Draw a border around the track preview
            .BackColor = objRace.TRACK_COLOUR 'Set the colour according to the track colour
            .ForeColor = objOption.COL_TEXT
        End With
        .cmdS5.caption = GetText(g_arr_Text, "START014") 'button "Speed and form"
        .cmdS6.caption = GetText(g_arr_Text, "START015") 'button "Odds"
        'Height of the pop-up
        If objOption.BET_MODE = True And objRace.BETTING_ALLOWED = "Y" Then
            .Height = 410 'If the betting mode is enabled
        Else
            .Height = 280 'If the betting mode is disabled
        End If
        'Width of the pop-up
        .width = 550
        .show (vbModal)
    End With
End Sub

Public Sub GetNumberBetSlips()
    'Count and display the number of bet slips submitted
    frmStart.lblBet02.caption = GetText(g_arr_Text, "START012") & ": " & g_colBetSlips.count
End Sub

Public Sub Gambler()
    'Pop-up for the name of the gambler who is placing a bet
    Call ShowInputPopup(objRace.RACE_NAME & " " & objRace.RACE_YEAR, _
        GetText(g_arr_Text, "BET002"), 120, 26, enumButton.CancelOK, vbModal)
    'Evaluate the input value
    If g_enumButton = enumButton.OK And Trim(g_strInpBoxReturnValue) <> "" Then _
        Call ShowBettingSlip(g_strInpBoxReturnValue) 'Show the pop-up with the betting slip
End Sub

'Provide information for the betting slip
Private Sub ShowBettingSlip(strName As String)
    With frmBetSlip
        .caption = strName
        .lblC1 = objRace.TRACK_NAME & " - " & objRace.TRACK_LOCATION & " (" & objRace.COUNTRY_NAME & ")"
        .lblC2 = objRace.RACE_NAME & " " & objRace.RACE_YEAR
        .Height = 334
        .width = 912
        .show (vbModal)
    End With
End Sub

'Pop-up with speed bars and odds
Public Sub ShowSpeed(ODDS As Boolean)
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Dim min As Long, max As Long 'Min/max values of the payout on win bet
    Dim group As Integer 'Counter for label groups
    Dim i As Long, j As Long
    
    'Texts for existing labels
    With frmOdds
        .caption = objRace.RACE_NAME & " " & objRace.RACE_YEAR
        .width = 560
        .Height = 80
        .lblO0.caption = GetText(g_arr_Text, "ODDS001") '"No."
        .lblO1.caption = GetText(g_arr_Text, "ODDS002") '"Name"
        If ODDS Then
            With .lblO2
                .caption = GetText(g_arr_Text, "ODDS003") '"Odds"
                .ControlTipText = GetText(g_arr_Text, "ODDS004")
                .TextAlign = fmTextAlignRight
            End With
        Else
            With .lblO2
                .caption = ""
                .ControlTipText = ""
            End With
        End If
    End With
    
    group = 1

    'Find the minimum and maximum payout
    For i = 1 To UBound(g_arr_varHorses)
        If min = 0 Or g_arr_varHorses(i, 17) < min Then min = g_arr_varHorses(i, 17)
        If g_arr_varHorses(i, 17) > max Then max = g_arr_varHorses(i, 17)
    Next i
    
    'Display the horses sorted by odds in ascending order
    For i = min To max
        For j = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(j, 17) = i Then
            
                'Create a label with the horse name and number
                Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True)
                With g_objLabel
                    .Font.name = "Tahoma"
                    .Font.size = 12
                    .left = 12
                    .top = 28 + group * 18
                    .width = 200
                    .TextAlign = fmTextAlignLeft
                    .caption = "#" & g_arr_varHorses(j, 11) & vbTab & g_arr_varHorses(j, 1)
                    If g_arr_varHorses(j, 0) <> "START" Then .Font.Strikethrough = True
                End With
                
                'Adjust the height of the pop-up
                frmOdds.Height = frmOdds.Height + g_objLabel.Height
                
                'In case of displaying odds: create a label for payout
                If ODDS Then
                    Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True)
                    With g_objLabel
                        .Font.name = "Tahoma"
                        .Font.size = 12
                        .left = 220
                        .top = 28 + group * 18
                        .width = 62
                        .TextAlign = fmTextAlignRight
                        .caption = g_arr_varHorses(j, 17) & ":10"
                        If g_arr_varHorses(j, 0) <> "START" Then .Font.Strikethrough = True
                    End With
                End If

                'Create a label for the upper horizontal bar (basic speed)
                Dim xxx As Integer 'ToDo Marco --> irgendwie auslagern in TEC 2 oder okay hier??
                Select Case objRace.RACE_ID
                    Case "SPACE"
                        xxx = 100
                    Case Else
                        xxx = 1500
                End Select
                
                If g_arr_varHorses(j, 0) = "START" Then
                    Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True)
                    With g_objLabel
                        .left = 290
                        .top = 27 + group * 18
                        .Height = 7
                        .width = 150 + ((g_arr_varHorses(j, 5) - xxx) * 5)
                        .BackColor = 14395790 'Blue bar
                        #If Debugging Then 'Description inside the bar for the basic speed
                            .caption = "Speed: " & g_arr_varHorses(j, 5) _
                                & " >> " & 150 + ((g_arr_varHorses(j, 5) - xxx) * 5) _
                                & " px (150 + " & ((g_arr_varHorses(j, 5) - xxx) * 5) & ")"
                            .Font.size = 6
                        #End If
                    End With
                End If
                
                'Create a label for the lower horizontal bar (estimated horse condition)
                If g_arr_varHorses(j, 0) = "START" Then
                    Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True)
                    With g_objLabel
                        .left = 290
                        .top = 27 + group * 18 + 8
                        .Height = 7
                        .width = 150 + ((g_arr_varHorses(j, 6) - xxx) * 5) + g_arr_varHorses(j, 18)
                        .BackColor = 6740479 'Yellow bar
                        #If Debugging Then 'Description inside the bar for the daily form
                            .caption = "Cond : " & g_arr_varHorses(j, 6) _
                                & " >> " & 150 + ((g_arr_varHorses(j, 6) - xxx) * 5) + g_arr_varHorses(j, 18) _
                                & " px (150 + " & ((g_arr_varHorses(j, 6) - xxx) * 5) _
                                & " + " & g_arr_varHorses(j, 18) & " (est. error)"
                            .Font.size = 6
                        #End If
                    End With
                End If
                
                group = group + 1 'Next section
            End If
        Next j
    Next i

    'Create a label for the upper horizontal bar
    'which serves as a headline with description
    Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True)
    With g_objLabel
        With .Font
                .name = "Tahoma"
                .Bold = True
                .size = 10
        End With
        .left = 290
        .top = 6
        .Height = 15
        .width = 246
        .TextAlign = fmTextAlignCenter
        .BackColor = 14395790 'Blue
        .caption = GetText(g_arr_Text, "ODDS005") '"Basic speed"
        .ControlTipText = GetText(g_arr_Text, "ODDS006") & " " & g_arr_Grammar(5)
    End With

    'Create a label for the lower horizontal bar
    'which serves as a headline with description
    Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True)
    With g_objLabel
        With .Font
                .name = "Tahoma"
                .Bold = True
                .size = 10
        End With
        .left = 290
        .top = 22
        .Height = 15
        .width = 246
        .TextAlign = fmTextAlignCenter
        .BackColor = 6740479 'Yellow
        .caption = GetText(g_arr_Text, "ODDS007") '"Form on the day - impression during warm-up"
        .ControlTipText = GetText(g_arr_Text, "ODDS008") & " " & g_arr_Grammar(5)
    End With

    frmOdds.show (vbModal) 'Show the pop-up
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "ShowSpeed(Odds As Boolean)")
    Call basAuxiliary.CodeCrash
End Sub

'Pop-up with a betting receipt
Public Sub ShowReceipt(id As Integer)
    Dim BET As String, horsename As String
    Dim i As Integer, j As Integer
    Dim bt() As Integer
    bt() = g_colBetSlips(id).BET 'Get the array with the guesses
    For i = 1 To UBound(bt) 'Loop through the array
        For j = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(j, 11) = bt(i) Then 'Find the horse name
                horsename = g_arr_varHorses(j, 1)
                Exit For
            End If
        Next j
        BET = BET & bt(i) & " " & horsename & vbNewLine 'Horse number and name
    Next i

    'Write all data to the receipt
    With frmReceipt
        .caption = g_colBetSlips(id).GAMBLERNAME
        .lblR1 = UCase(objRace.TRACK_LOCATION) ' & " (" & objRace.COUNTRY_NAME & ")")
                                            '(comment in for adding the country)
        .lblR2 = UCase(objRace.RACE_NAME)
        .lblR3 = objRace.NUMBER_STARTING & " " & UCase(g_arr_Grammar(4))
        .lblR4 = UCase(g_colBetSlips(id).BET_TYPE_TEXT)
        .lblR5 = UCase(BET)
        .lblR6 = UCase(GetText(g_arr_Text, "BET036") & " " & Format(g_colBetSlips(id).STAKE, "0.00") & " " & GetText(g_arr_Text, "BET035"))
        .lblR7 = g_colBetSlips(id).BET_ID
        .Height = 280
        .width = 215
        .show (vbModal)
    End With
End Sub


'CALLBACKS (Excel ribbon events)
'-------------------------------

'AI mode (Callback for Excel ribbon): customUI.onLoad
Private Sub AI_GaloppSimAddinInitialize(ribbon As IRibbonUI)
    Set g_RibbonGaloppSim = ribbon
    If g_wksTEXT Is Nothing Then Set g_wksTEXT = Table_TEXT
    Call GetTextComponents
End Sub

'AI mode (Callback for Excel ribbon): Labels
Private Sub AI_GetLabel(control As IRibbonControl, ByRef returnedVal)
    
    'Display the start screen in AI edition when activating the GALOPPSIM menu tab the first time
    If AI_started = False Then
        AI_started = True
        Call TitleScreen
    End If
    
    Select Case control.id
        Case "group01GALOPPSIM" 'Group "Settings"
            returnedVal = GetText(g_arr_Text, "AI001")
            Case "btn10GALOPPSIM" 'Button "Race options"
                returnedVal = GetText(g_arr_Text, "BTN001")
            Case "menu02GALOPPSIM" 'Menu "Excel options"
                returnedVal = GetText(g_arr_Text, "BTN002")
                Case "cb01aGALOPPSIM" 'Button "Excel mode"
                    returnedVal = GetText(g_arr_Text, "EXCELOPT010")
                Case "cb01bGALOPPSIM" 'Button "TV mode with Excel menu strip"
                    returnedVal = GetText(g_arr_Text, "EXCELOPT011")
                Case "cb01cGALOPPSIM" 'Button "TV mode (full-screen)"
                    returnedVal = GetText(g_arr_Text, "EXCELOPT012")
                
        Case "group02GALOPPSIM" 'Group "Race"
            returnedVal = GetText(g_arr_Text, "AI002")
            Case "btn30GALOPPSIM" 'Button "Start race"
                returnedVal = GetText(g_arr_Text, getCaptionStartBtn(objOption.BET_MODE))
            Case "combo01InstalledRaces" 'ComboBox "Race selection"
                returnedVal = GetText(g_arr_Text, "RACE003")
            Case "btn31GALOPPSIM" 'Button "Photo of the finish"
                returnedVal = GetText(g_arr_Text, "BTN004")
            Case "btn32GALOPPSIM" 'Button "Ranking list"
                returnedVal = GetText(g_arr_Text, "BTN005")
            Case "btn33GALOPPSIM" 'Button "Winner"
                returnedVal = GetText(g_arr_Text, "BTN006")
            Case "btn34GALOPPSIM" 'Button "Betting analysis"
                returnedVal = GetText(g_arr_Text, "BTN007")
            
        Case "menu01GALOPPSIM" 'Menu "Language"
            returnedVal = GetText(g_arr_Text, "LANGUAGE001")
            Case "btn01aGALOPPSIM" 'Button "Deutsch"
                returnedVal = GetText(g_arr_Text, "LANGUAGE002")
            Case "btn01bGALOPPSIM" 'Button "English"
                returnedVal = GetText(g_arr_Text, "LANGUAGE003")
            Case "btn01eGALOPPSIM" 'Button "Bulgarian"
                returnedVal = GetText(g_arr_Text, "LANGUAGE006")
        Case "btn20GALOPPSIM" 'Button "Colour mode"
            returnedVal = GetText(g_arr_Text, "BTN021")
        Case "btn40GALOPPSIM" 'Button "Info"
            returnedVal = GetText(g_arr_Text, "BTN009")
        Case "btn50GALOPPSIM" 'Button "Warning"
            returnedVal = GetText(g_arr_Text, "BTN010")
        Case "btn60GALOPPSIM" 'Button "Movie"
            returnedVal = GetText(g_arr_Text, "BTN011")
        Case "btn05GALOPPSIM" 'Button "Title screen"
            returnedVal = GetText(g_arr_Text, "BTN020")
        Case "btn70GALOPPSIM" 'Button "Termination"
            returnedVal = GetText(g_arr_Text, "BTN012")
            
        Case "group04GALOPPSIM" 'Group "Replay"
            returnedVal = GetText(g_arr_Text, "AI003")
        Case "btn80GALOPPSIM" 'Button "Race Replay"
            returnedVal = GetText(g_arr_Text, "BTN025")
        Case "btn81GALOPPSIM" 'Button "Save Race"
            returnedVal = GetText(g_arr_Text, "BTN026")
        Case "btn82GALOPPSIM" 'Button "Load Race"
            returnedVal = GetText(g_arr_Text, "BTN027")
    End Select
End Sub

'AI mode (Callback for Excel ribbon): Tooltips (Screentips)
Private Sub AI_GetScreentip(control As IRibbonControl, ByRef screentip)
    Select Case control.id
        Case "btn10GALOPPSIM" 'Button "Race options"
            screentip = GetText(g_arr_Text, "BTN001")
        Case "menu02GALOPPSIM" 'Menu "Excel options"
            screentip = GetText(g_arr_Text, "BTN002")
        Case "cb01aGALOPPSIM" 'Button "Excel mode"
            screentip = GetText(g_arr_Text, "EXCELOPT010")
        Case "cb01bGALOPPSIM" 'Button "TV mode with Excel menu strip"
            screentip = GetText(g_arr_Text, "EXCELOPT011")
        Case "cb01cGALOPPSIM" 'Button "TV mode (full-screen)"
            screentip = GetText(g_arr_Text, "EXCELOPT012")
        Case "btn30GALOPPSIM" 'Button "Start race"
            screentip = GetText(g_arr_Text, getCaptionStartBtn(objOption.BET_MODE))
        Case "combo01InstalledRaces" 'ComboBox "Race selection"
            screentip = GetText(g_arr_Text, "TIP024")
        Case "btn31GALOPPSIM" 'Button "Photo of the finish"
            screentip = GetText(g_arr_Text, "BTN004")
        Case "btn32GALOPPSIM" 'Button "Ranking list"
            screentip = GetText(g_arr_Text, "BTN005")
        Case "btn33GALOPPSIM" 'Button "Winner"
            screentip = GetText(g_arr_Text, "BTN006")
        Case "btn34GALOPPSIM" 'Button "Betting analysis"
            screentip = GetText(g_arr_Text, "BTN007")
        Case "menu01GALOPPSIM" 'Menu "Language"
            screentip = GetText(g_arr_Text, "LANGUAGE001")
        Case "btn20GALOPPSIM" 'Button "Colour mode"
            screentip = GetText(g_arr_Text, "BTN021")
        Case "btn40GALOPPSIM" 'Button "Info"
            screentip = GetText(g_arr_Text, "TIP026")
        Case "btn50GALOPPSIM" 'Button "Warning"
            screentip = GetText(g_arr_Text, "BTN010")
        Case "btn60GALOPPSIM" 'Button "Movie"
            screentip = GetText(g_arr_Text, "BTN030")
        Case "btn05GALOPPSIM" 'Button "Title screen"
            screentip = GetText(g_arr_Text, "BTN020")
        Case "btn70GALOPPSIM" 'Button "Termination"
            screentip = GetText(g_arr_Text, "TIP039")
        Case "btn80GALOPPSIM" 'Button "Race Replay"
            screentip = GetText(g_arr_Text, "BTN025")
        Case "btn81GALOPPSIM" 'Button "Save Race"
            screentip = GetText(g_arr_Text, "BTN023")
        Case "btn82GALOPPSIM" 'Button "Load Race"
            screentip = GetText(g_arr_Text, "BTN024")
    End Select
End Sub

'AI mode (Callback for Excel ribbon): Tooltips (Supertips)
Private Sub AI_GetSupertip(control As IRibbonControl, ByRef supertip)
    Select Case control.id
        Case "btn10GALOPPSIM" 'Button "Race options"
            supertip = GetText(g_arr_Text, "TIP037")
        Case "menu02GALOPPSIM" 'Menu "Excel options"
            supertip = GetText(g_arr_Text, "TIP045")
        Case "cb01aGALOPPSIM" 'Button "Excel mode"
            supertip = GetText(g_arr_Text, "TIP013")
        Case "cb01bGALOPPSIM" 'Button "TV mode with Excel menu strip"
            supertip = GetText(g_arr_Text, "TIP014")
        Case "cb01cGALOPPSIM" 'Button "TV mode (full-screen)"
            supertip = GetText(g_arr_Text, "TIP015")
        Case "btn30GALOPPSIM" 'Button "Start race"
            supertip = GetText(g_arr_Text, "TIP046")
        Case "combo01InstalledRaces" 'ComboBox "Race selection"
            supertip = GetText(g_arr_Text, "TIP047")
        Case "btn31GALOPPSIM" 'Button "Photo of the finish"
            supertip = GetText(g_arr_Text, "TIP041")
        Case "btn32GALOPPSIM" 'Button "Ranking list"
            supertip = GetText(g_arr_Text, "TIP042")
        Case "btn33GALOPPSIM" 'Button "Winner"
            supertip = GetText(g_arr_Text, "TIP043")
        Case "btn34GALOPPSIM" 'Button "Betting analysis"
            supertip = GetText(g_arr_Text, "TIP044")
        Case "menu01GALOPPSIM" 'Menu "Language"
            supertip = GetText(g_arr_Text, "TIP048")
        Case "btn20GALOPPSIM" 'Button "Colour mode"
            supertip = GetText(g_arr_Text, "TIP038")
        Case "btn40GALOPPSIM" 'Button "Info"
            supertip = GetText(g_arr_Text, "TIP027")
        Case "btn50GALOPPSIM" 'Button "Warning"
            supertip = GetText(g_arr_Text, "TIP049")
        Case "btn60GALOPPSIM" 'Button "Movie"
            supertip = GetText(g_arr_Text, "BTN031")
        Case "btn05GALOPPSIM" 'Button "Title screen"
            supertip = GetText(g_arr_Text, "TIP012")
        Case "btn70GALOPPSIM" 'Button "Termination"
            supertip = GetText(g_arr_Text, "TIP040")
        Case "btn80GALOPPSIM" 'Button "Race Replay"
            supertip = GetText(g_arr_Text, "TIP057")
        Case "btn81GALOPPSIM" 'Button "Save Race"
            supertip = GetText(g_arr_Text, "TIP058")
        Case "btn82GALOPPSIM" 'Button "Load Race"
            supertip = GetText(g_arr_Text, "TIP059")
    End Select
End Sub

'AI mode (Callback for Excel ribbon): Status of buttons (enabled or disabled)
Private Sub AI_IsButtonEnabled(control As IRibbonControl, ByRef returnedVal)
    Select Case control.id
        Case "btn31GALOPPSIM" 'Button "Photo of the finish"
            returnedVal = objRace.STARTED
        Case "btn32GALOPPSIM" 'Button "Ranking list"
            returnedVal = objRace.STARTED
        Case "btn33GALOPPSIM" 'Button "Photo of the winner"
            returnedVal = objRace.STARTED
        Case "btn34GALOPPSIM" 'Button "Betting analysis"
            returnedVal = objRace.STARTED And objOption.BET_PLACED
        Case "btn80GALOPPSIM" 'Button "Race Replay"
            returnedVal = (objRace.STARTED Or objRace.LOADED)
        Case "btn81GALOPPSIM" 'Button "Save Race"
            returnedVal = objRace.STARTED And Not objRace.LOADED = True
    End Select
End Sub

'AI mode (Callback for Excel ribbon): Get values of checkboxes in the Excel ribbon (getPressed)
Public Sub AI_ExcelModeGet(control As IRibbonControl, ByRef standardwert)
    Select Case control.id
        Case "cb01aGALOPPSIM" 'Excel mode
            standardwert = (objOption.EXCEL_MODE = "normal")
        Case "cb01bGALOPPSIM" 'TV mode with Excel menu strip
            standardwert = (objOption.EXCEL_MODE = "TVmenu")
        Case "cb01cGALOPPSIM" 'TV mode (full-screen)
            standardwert = (objOption.EXCEL_MODE = "TVfull")
    End Select
End Sub

'AI mode (Callback for Excel ribbon): Set values of checkboxes in the Excel ribbon (onAction)
Public Sub AI_ExcelModeSet(control As IRibbonControl, pressed As Boolean)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Select Case control.id
        Case "cb01aGALOPPSIM" 'Excel mode
            objOption.EXCEL_MODE = "normal"
            Call ResetExcelOptions
        Case "cb01bGALOPPSIM" 'TV mode with Excel menu strip
            objOption.EXCEL_MODE = "TVmenu"
            Call ExcelOptionsTVmenu
        Case "cb01cGALOPPSIM" 'TV mode (full-screen)
            objOption.EXCEL_MODE = "TVfull"
            Call ExcelOptionsTVmenu
    End Select
    
    g_RibbonGaloppSim.Invalidate 'refresh the status of the checkboxes
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Betting" "Race options"
Private Sub AI_OptionsRace(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    Call GetAnimalGrammar
    frmOptionsRace.show (vbModal) 'display UserForm (modal)
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Start the race"
Private Sub AI_StartRace(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    'Leave the current race?
    If objRace.STARTED Then
        Dim strNextRace As String
        With ThisWorkbook.Worksheets(objRace.SELECTED)
            strNextRace = .Cells(basAuxiliary.GetRow(ThisWorkbook.Worksheets(objRace.SELECTED), "RACE NAME"), 2).Value & " " & _
                .Cells(basAuxiliary.GetRow(ThisWorkbook.Worksheets(objRace.SELECTED), "YEAR"), 2).Value & " (" & _
                .Cells(basAuxiliary.GetRow(ThisWorkbook.Worksheets(objRace.SELECTED), "DISTANCE METRES"), 2).Value & "m) - " & _
                .Cells(basAuxiliary.GetRow(ThisWorkbook.Worksheets(objRace.SELECTED), "TRACK LOCATION"), 2).Value
        End With

        'Pop-up
    Call ShowMessagePopup(g_c_tool, GetText(g_arr_Text, "RACE003") & ": " & strNextRace, _
        enumButton.CancelOK, vbModal)
            
            'Evaluate the return value
            If g_enumButton = enumButton.Cancel Then Exit Sub
            
            If g_strPlayMode = "AI" Then g_RibbonGaloppSim.Invalidate 'reset Excel ribbon
            Call ShowNewRaceScreen("NEWRACE2")
    End If

    Call NewRace(False)
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Ranking list"
Private Sub AI_Results(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call ShowRankingList
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Photo of the winner"
Private Sub AI_Winner(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call ShowWinnerPhoto(False, True)
End Sub

'Show the photo of the winner
Public Sub ShowWinnerPhoto(Optional test As Boolean, Optional zoom As Boolean)
    If objRace.STARTED Then
        Call basAuxiliary.ActivateRaceSheet
        Call DrawWinnerPhoto(zoom)
        If g_strPlayMode = "RS" And Not test Then frmRS_navigation.show (vbModeless)
    End If
End Sub

'Show the ranking list with the race results
Public Sub ShowRankingList(Optional test As Boolean)
    If objRace.STARTED Then
        Call basAuxiliary.ActivateRaceSheet
        Call FormatPhotoAndRanking
        Call basAuxiliary.Scroll(objRace.METRES + objBasicData.LEFT_COLS + 15 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR), objBasicData.TOP_ROWS + (objRace.NUMBER_ENROLLED * 2 + 9))
        Call DrawRankingList(False, test, False)
        If g_strPlayMode = "RS" And Not test Then frmRS_navigation.show (vbModeless)
    End If
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Photo of the finish"
Private Sub AI_FinishPhoto(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call ShowFinishPhoto
End Sub

'Show the photo of the finish
Public Sub ShowFinishPhoto(Optional test As Boolean)
    If objRace.STARTED Then
        Call basAuxiliary.ActivateRaceSheet
        Call FormatPhotoAndRanking
        Call basAuxiliary.Scroll(objRace.METRES + objBasicData.LEFT_COLS + 15 + objBasicData.AFTER_FIN_COLS + (2 * objOption.SPEED_FACTOR), objBasicData.TOP_ROWS + 1)
        Call DrawFinishPhoto
        If g_strPlayMode = "RS" And Not test Then frmRS_navigation.show (vbModeless)
    End If
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Betting analysis"
Private Sub AI_Betting(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call ShowBets
End Sub

'Show a pop-up with the betting analysis
Public Sub ShowBets()
    If objRace.STARTED And objOption.BET_PLACED Then
        Call AnalyseBettings
    End If
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Info section"
Private Sub AI_Info(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call ShowInfo
End Sub

'AI mode: Adapt Excel settings according to the selected mode
Public Sub AI_ExcelModeStart()
    Select Case objOption.EXCEL_MODE
        Case "normal"
            Call ResetExcelOptions
        Case "TVmenu"
            Call ExcelOptionsTVmenu
        Case "TVfull"
            Call ExcelOptionsTVfull
    End Select
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Colour mode"
Private Sub AI_ColourMode(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call ColourModeSelection
End Sub

'AI mode: Reset Excel settings
Public Sub AI_ExcelModeEnd()
    Select Case objOption.EXCEL_MODE
        Case "normal"
            
        Case "TVmenu"
            
        Case "TVfull"
            Call ExcelOptionsTVmenu
    End Select
End Sub

'Provide content for the pop-up with the tool information
Private Sub ShowInfo()
    
    With frmInfo
        'Version and license
        .caption = g_c_tool & " - " & GetText(g_arr_Text, "INFO08")
        .lbl_info01.caption = GetText(g_arr_Text, "GEN01") & vbNewLine & GetText(g_arr_Text, "GEN02")
        
        For i = 0 To 6 'Captions of the tabs
            .multiPage_info.Pages(i).caption = GetText(g_arr_Text, "PAGE0" & i + 1)
        Next i
            .multiPage_info.Value = 0 'Set the focus on the first tab
            
        'Tab "GaloppSim"
        '---------------
        With .lbl_info_galoppsim01
            .caption = GetText(g_arr_Text, "INFO01") & vbNewLine & vbNewLine _
                    & GetText(g_arr_Text, "INFO02") & vbNewLine & vbNewLine _
                    & GetText(g_arr_Text, "INFO03") & vbNewLine & vbNewLine _
                    & GetText(g_arr_Text, "INFO04") & vbNewLine & vbNewLine _
                    & GetText(g_arr_Text, "INFO05") & vbNewLine & vbNewLine _
                    & GetText(g_arr_Text, "INFO06") & vbNewLine & vbNewLine _
                    & GetText(g_arr_Text, "INFO07") & vbNewLine & vbNewLine
            .width = 460 'Fixed label width
            .AutoSize = True 'Label height depending on the content
        End With
        With .multiPage_info.Pages(0)
            .ScrollBars = fmScrollBarsVertical 'Vertical scrollbar
            .ScrollHeight = .lbl_info_galoppsim01.Height 'Height of the vertical scrolling
            .KeepScrollBarsVisible = fmScrollBarsNone 'Show the scrollbar only if needed
        End With

        'Tab "Team"
        '----------
        'Marco Matjes
        .lbl_info_team01a.caption = GetText(g_arr_Text, "TEAM001")
        .lbl_info_team01b.caption = GetText(g_arr_Text, "TEAM002")
        .img_info_team01.ControlTipText = GetText(g_arr_Text, "TEAM003")
        'Florian
        .lbl_info_team02a.caption = GetText(g_arr_Text, "TEAM004")
        .lbl_info_team02b.caption = GetText(g_arr_Text, "TEAM005")
        .img_info_team02.ControlTipText = GetText(g_arr_Text, "TEAM006")
        'Paul
        .lbl_info_team03a.caption = GetText(g_arr_Text, "TEAM007")
        .lbl_info_team03b.caption = GetText(g_arr_Text, "TEAM008")
        .img_info_team03.ControlTipText = GetText(g_arr_Text, "TEAM009")
        'Michael
        .lbl_info_team04a.caption = GetText(g_arr_Text, "TEAM010")
        .lbl_info_team04b.caption = GetText(g_arr_Text, "TEAM011")
        .img_info_team04.ControlTipText = GetText(g_arr_Text, "TEAM012")
        'Meike
        .lbl_info_team05a.caption = GetText(g_arr_Text, "TEAM013")
        .lbl_info_team05b.caption = GetText(g_arr_Text, "TEAM014")
        .img_info_team05.ControlTipText = GetText(g_arr_Text, "TEAM015")
        'Atanas
        .lbl_info_team07a.caption = GetText(g_arr_Text, "TEAM019")
        .lbl_info_team07b.caption = GetText(g_arr_Text, "TEAM020")
        .img_info_team07.ControlTipText = GetText(g_arr_Text, "TEAM021")
        'Enno
        .lbl_info_team09a.caption = GetText(g_arr_Text, "TEAM025")
        .lbl_info_team09b.caption = GetText(g_arr_Text, "TEAM026")
        .img_info_team09.ControlTipText = GetText(g_arr_Text, "TEAM027")
        'Farida
        .lbl_info_team10a.caption = GetText(g_arr_Text, "TEAM028")
        .lbl_info_team10b.caption = GetText(g_arr_Text, "TEAM029")
        .img_info_team10.ControlTipText = GetText(g_arr_Text, "TEAM030")
        
        'Vertical scrollbar
        With .multiPage_info.Pages(1)
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = 450
            .KeepScrollBarsVisible = fmScrollBarsNone
        End With
            
        'Tab "Algorithms"
        '----------------
        For i = 0 To 8 'Captions of the tabs
            .multiPage_algo.Pages(i).caption = GetText(g_arr_Text, "PAGEALGO0" & i + 1)
        Next i
        .multiPage_algo.MultiRow = True 'Display all tabs in rows without scrolling
        .multiPage_algo.Value = 0 'Set the focus on the first tab
        
        '"It's complex implementation."
        .img_info_algorithms01.ControlTipText = GetText(g_arr_Text, "ALGO01")
        .img_info_algorithms02.ControlTipText = GetText(g_arr_Text, "ALGO01")
        
        'Checkbox "Stop algorithms!"
        With .chk_info_algorithms01
            .caption = GetText(g_arr_Text, "ALGO02")
            .Font.size = 20
            .Font.Bold = True
            .ControlTipText = GetText(g_arr_Text, "ALGO03")
        End With
        
        'Algorithm 01 'Overall race algorithm
        Call LabelAlgo(.lbl_algo_01_00, "PAGEALGO01", 6, 6, 330, True, True)
        Call LabelAlgo(.lbl_algo_01_01, "ALGO10", frmInfo.lbl_algo_01_00.Height + 12, 6, 330, True)
        Call LabelAlgo(.lbl_algo_01_02, "ALGO11", frmInfo.lbl_algo_01_00.Height + frmInfo.lbl_algo_01_01.Height + 24, 6, 330, True)
        With .multiPage_algo(0)
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = .lbl_algo_01_00.Height + .lbl_algo_01_01.Height + .lbl_algo_01_02.Height + 30
            .KeepScrollBarsVisible = fmScrollBarsNone
        End With
            
        'Algorithm 02 'Form on the day algorithm
        Call LabelAlgo(.lbl_algo_02_00, "PAGEALGO02", 6, 6, 330, True, True)
        Call LabelAlgo(.lbl_algo_02_01, "ALGO15", frmInfo.lbl_algo_02_00.Height + 12, 6, 330, True)
        Call LabelAlgo(.lbl_algo_02_02, "ALGO16", frmInfo.lbl_algo_02_00.Height + frmInfo.lbl_algo_02_01.Height + 24, 6, 330, True)
        With .multiPage_algo(1)
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = .lbl_algo_02_00.Height + .lbl_algo_02_01.Height + .lbl_algo_02_02.Height + 30
            .KeepScrollBarsVisible = fmScrollBarsNone
        End With
            
        'Algorithm 03 'Loop algorithm
        Call LabelAlgo(.lbl_algo_03_00, "PAGEALGO03", 6, 6, 330, True, True)
        Call LabelAlgo(.lbl_algo_03_01, "ALGO20", frmInfo.lbl_algo_03_00.Height + 12, 6, 330, True)
        Call LabelAlgo(.lbl_algo_03_02, "ALGO21", frmInfo.lbl_algo_03_00.Height + frmInfo.lbl_algo_03_01.Height + 24, 6, 330, True)
        With .multiPage_algo(2)
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = .lbl_algo_03_00.Height + .lbl_algo_03_01.Height + .lbl_algo_03_02.Height + 30
            .KeepScrollBarsVisible = fmScrollBarsNone
        End With
            
        'Algorithm 04 'Slipstream algorithm
         Call LabelAlgo(.lbl_algo_04_00, "PAGEALGO04", 6, 6, 330, True, True)
         Call LabelAlgo(.lbl_algo_04_01, "ALGO25", frmInfo.lbl_algo_04_00.Height + 12, 6, 330, True)
         Call LabelAlgo(.lbl_algo_04_02, "ALGO26", frmInfo.lbl_algo_04_00.Height + frmInfo.lbl_algo_04_01.Height + 24, 6, 330, True)
        With .multiPage_algo(3)
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = .lbl_algo_04_00.Height + .lbl_algo_04_01.Height + .lbl_algo_04_02.Height + 3
            .KeepScrollBarsVisible = fmScrollBarsNone
        End With
            
        'Algorithm 05 'Favourites calculation
        Call LabelAlgo(.lbl_algo_05_00, "PAGEALGO05", 6, 6, 330, True, True)
        Call LabelAlgo(.lbl_algo_05_01, "ALGO30", frmInfo.lbl_algo_05_00.Height + 12, 6, 330, True)
        Call LabelAlgo(.lbl_algo_05_02, "ALGO31", frmInfo.lbl_algo_05_00.Height + frmInfo.lbl_algo_05_01.Height + 24, 6, 330, True)
        With .multiPage_algo(4)
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = .lbl_algo_05_00.Height + .lbl_algo_05_01.Height + .lbl_algo_05_02.Height + 30
            .KeepScrollBarsVisible = fmScrollBarsNone
        End With
            
        'Algorithm 06 'Warm-up impression algorithm
        Call LabelAlgo(.lbl_algo_06_00, "PAGEALGO06", 6, 6, 330, True, True)
        Call LabelAlgo(.lbl_algo_06_01, "ALGO35", frmInfo.lbl_algo_06_00.Height + 12, 6, 330, True)
        Call LabelAlgo(.lbl_algo_06_02, "ALGO36", frmInfo.lbl_algo_06_00.Height + frmInfo.lbl_algo_06_01.Height + 24, 6, 330, True)
        With .multiPage_algo(5)
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = .lbl_algo_06_00.Height + .lbl_algo_06_01.Height + .lbl_algo_06_02.Height + 30
            .KeepScrollBarsVisible = fmScrollBarsNone
        End With
            
        'Algorithm 07 'Betting odds algorithm
        Call LabelAlgo(.lbl_algo_07_00, "PAGEALGO07", 6, 6, 330, True, True)
        Call LabelAlgo(.lbl_algo_07_01, "ALGO40", frmInfo.lbl_algo_07_00.Height + 12, 6, 330, True)
        Call LabelAlgo(.lbl_algo_07_02, "ALGO41", frmInfo.lbl_algo_07_00.Height + frmInfo.lbl_algo_07_01.Height + 24, 6, 330, True)
        With .multiPage_algo(6)
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = .lbl_algo_07_00.Height + .lbl_algo_07_01.Height + .lbl_algo_07_02.Height + 30
            .KeepScrollBarsVisible = fmScrollBarsNone
        End With
            
        'Algorithm 08 'Splashwater algorithm
        Call LabelAlgo(.lbl_algo_08_00, "PAGEALGO08", 6, 6, 330, True, True)
        Call LabelAlgo(.lbl_algo_08_01, "ALGO42", frmInfo.lbl_algo_08_00.Height + 12, 6, 330, True)
        Call LabelAlgo(.lbl_algo_08_02, "ALGO43", frmInfo.lbl_algo_08_00.Height + frmInfo.lbl_algo_08_01.Height + 24, 6, 330, True)
        With .multiPage_algo(7)
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = .lbl_algo_08_00.Height + .lbl_algo_08_01.Height + .lbl_algo_08_02.Height + 30
            .KeepScrollBarsVisible = fmScrollBarsNone
        End With
            
        'Algorithm 09 'Colourgrey algorithm
        Call LabelAlgo(.lbl_algo_09_00, "PAGEALGO09", 6, 6, 330, True, True)
        Call LabelAlgo(.lbl_algo_09_01, "ALGO44", frmInfo.lbl_algo_09_00.Height + 12, 6, 330, True)
        Call LabelAlgo(.lbl_algo_09_02, "ALGO45", frmInfo.lbl_algo_09_00.Height + frmInfo.lbl_algo_09_01.Height + 24, 6, 330, True)
        With .multiPage_algo(8)
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = .lbl_algo_09_00.Height + .lbl_algo_09_01.Height + .lbl_algo_09_02.Height + 30
            .KeepScrollBarsVisible = fmScrollBarsNone
            End With
                
        'Tab "Code"
        With .lbl_info_code01
            .caption = GetText(g_arr_Text, "CODE01")
            .Font.size = 12
            .WordWrap = True
            .AutoSize = True
        End With
        .btn_info_code01.ControlTipText = GetText(g_arr_Text, "CODE02")
        With .lbl_info_code03
            .caption = GetText(g_arr_Text, "CODE03")
            .Font.size = 12
            .WordWrap = True
            .AutoSize = True
        End With
        With .lbl_info_code04
            .caption = GetText(g_arr_Text, "CODE04") & vbNewLine & vbNewLine _
                        & GetText(g_arr_Text, "CODE05")
            .WordWrap = True
            .AutoSize = True
        End With
        With .multiPage_info.Pages(3)
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = .lbl_info_code04.Height
            .KeepScrollBarsVisible = fmScrollBarsNone
        End With
        
        'Tab "Donation"
        With .lbl_info_donation01
            .Font.size = 12
            .caption = GetText(g_arr_Text, "DON01") & vbNewLine & vbNewLine _
                        & GetText(g_arr_Text, "DON02")
            .AutoSize = True
        End With
        .btn_info_donation01.ControlTipText = GetText(g_arr_Text, "DON03")
        With .btn_info_donation02
            .caption = GetText(g_arr_Text, "DON04")
            .Font.size = 24
            .ControlTipText = GetText(g_arr_Text, "DON05")
        End With

        'Tab "Contact & Social media"
        With .lbl_info_contact01a
            .caption = GetText(g_arr_Text, "CON01") & vbNewLine _
                        & GetText(g_arr_Text, "CON02")
            .WordWrap = True
        End With
        .lbl_info_contact01b.caption = GetText(g_arr_Text, "CON03a")
        .lbl_info_contact01c.caption = GetText(g_arr_Text, "CON03b")
        With .btn_info_contact01
            .caption = GetText(g_arr_Text, "CON04")
            .ControlTipText = GetText(g_arr_Text, "CON05")
            .WordWrap = True
        End With
        .btn_info_contact02.ControlTipText = GetText(g_arr_Text, "CON06")
        With .lbl_info_contact02
            .Font.size = 12
            .TextAlign = fmTextAlignRight
            .caption = GetText(g_arr_Text, "CON07")
        End With
        With .btn_info_contact03
            .caption = GetText(g_arr_Text, "CON08")
            .ControlTipText = GetText(g_arr_Text, "CON09")
            .WordWrap = True
        End With
        With .btn_info_contact04
            .caption = GetText(g_arr_Text, "CON10")
            .ControlTipText = GetText(g_arr_Text, "CON11")
            .WordWrap = True
        End With
        With .btn_info_contact05
            .caption = GetText(g_arr_Text, "CON12")
            .ControlTipText = GetText(g_arr_Text, "CON13")
            .WordWrap = True
        End With
        With .lbl_info_contact03
            .ControlTipText = GetText(g_arr_Text, "CON14")
            .WordWrap = True
        End With

        'Tab "Privacy policy"
        With .lbl_info_privacy01
            .caption = GetText(g_arr_Text, "PRIVACY01") & " " _
                        & GetText(g_arr_Text, "PRIVACY02")
            .WordWrap = True
        End With
            
        'Show the pop-up
        .Height = 420
        .width = 523
        .show (vbModal)
    End With
End Sub

'Label for algorithm details
Private Sub LabelAlgo(lbl As Object, text As String, top As Integer, left As Integer, width As Integer, size As Boolean, Optional fb As Boolean)
    With lbl
        .caption = GetText(g_arr_Text, text)
        .top = top
        .left = left
        .width = width
        .Font.Bold = fb
        .AutoSize = size
    End With
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Warning notice"
Private Sub AI_Warning(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call ShowWarning
End Sub

'Show a pop-up with a warning message
Private Sub ShowWarning()
    Dim strWarningMessage As String
    strWarningMessage = GetText(g_arr_Text, "WARN001") & vbNewLine & GetText(g_arr_Text, "WARN002")
    Call ShowInfoPopup(GetText(g_arr_Text, "USERFORM003"), strWarningMessage, True, vbModal)
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Play movie"
Private Sub AI_Movie2017(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call GaloppSimMovie2017
End Sub

'Play the GaloppSim Movie (2017)
Private Sub GaloppSimMovie2017()
    'Close pop-ups if visible
    If frmBettingAnalysis.Visible Then Unload frmBettingAnalysis 'Betting analysis
    If frmRS_navigation.Visible Then Unload frmRS_navigation 'Navigation panel (RS edition only)
    'Play the movie
    Call basMovie2017.PlayMovie2017
End Sub

'Colour mode selection
Private Sub ColourModeSelection()
    'Close pop-ups if visible
    If frmBettingAnalysis.Visible Then Unload frmBettingAnalysis 'Betting analysis
    If frmRS_navigation.Visible Then Unload frmRS_navigation 'Navigation panel (RS edition only)
    Call basAuxiliary.Freeze(0, 0, False)
    frmColourMode.show
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Start screen"
Private Sub AI_Title(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call TitleScreen
        
End Sub

'Paint the GaloppSim title screen
'(currently used in AI edition only)
Public Sub TitleScreen()

    'Leave the current venue?
    If objRace.STARTED Then
        Call ShowMessagePopup(g_c_tool, GetText(g_arr_Text, "WARN003"), _
            enumButton.CancelOK, vbModal)
        'Evaluate the return value
        If g_enumButton = enumButton.Cancel Then Exit Sub
    End If

    Unload frmBettingAnalysis
    If objRace.STARTED Then objRace.STARTED = False
    
    'Reset the Excel ribbon
    g_RibbonGaloppSim.Invalidate

    Call CreateRaceSheet
    Call AI_ExcelModeStart
    Call AI_ExcelModeEnd
    
    With g_wksRace.Range(Cells(1, 1), Cells(40, 100))
        .ColumnWidth = ZoomLevelPictures()(0)
        .RowHeight = ZoomLevelPictures()(1)
    End With
    
    Call PaintPicture(g_wksPIC, g_wksRace, "AI_TITLE3", 100, 40, 1, 1)
    Call CursorAway

    objRace.STARTED = False
        
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Termination"
Private Sub AI_Close(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    If objRace.STARTED Then objRace.STARTED = False
    
    'Blatt l�schen
    Call AI_DeleteWorksheets
    Unload frmBettingAnalysis
    
    'Reset Excel ribbon
    g_RibbonGaloppSim.Invalidate

    'Reset Excel options
    Call ResetExcelOptions
    Application.ScreenUpdating = True
End Sub

'AI mode (Callback for Excel ribbon): Click on language button "DE"
Private Sub AI_LanguageDE(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    objOption.language = "DE"
    Call ChangeLanguage
End Sub

'AI mode (Callback for Excel ribbon): Click on language button "EN"
Private Sub AI_LanguageEN(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    objOption.language = "EN"
    Call ChangeLanguage
End Sub

'AI mode (Callback for Excel ribbon): Click on language button "BG"
Private Sub AI_LanguageBG(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    objOption.language = "BG"
    Call ChangeLanguage
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Replay"
Private Sub AI_Replay(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call RaceReplay
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Save race"
Private Sub AI_SaveRace(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call SaveRaceForReplay(False)
End Sub

'AI mode (Callback for Excel ribbon): Click on button "Load race"
Private Sub AI_LoadRace(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call LoadRaceForReplay
End Sub

'Change of the user interface language
Private Sub ChangeLanguage()
    Dim oleObj As OLEObject
    
    'Get text components of the selected language
    Call GetTextComponents
    Call GetAnimalGrammar
    
    If g_strPlayMode = "RS" Then
        'Loop through all runtime objects on the worksheet
        For Each oleObj In g_wksRace.OLEObjects
            If oleObj.name <> "CBraces" Then 'No change to the dropdown with the races
                Call RS_RefreshButtonTexts(oleObj.name) 'Refresh the button texts
            End If
        Next oleObj
    Else 'AI mode
        g_RibbonGaloppSim.Invalidate
    End If

End Sub

'Refresh the button label texts
Private Sub RS_RefreshButtonTexts(name As String)

    'Text for the start button depends on whether bettings are allowed
    Dim captionStart As String
    captionStart = basAuxiliary.getCaptionStartBtn(objOption.BET_MODE)
    
    Select Case name
        Case "startrace"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, captionStart)
        Case "finishphoto"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN004")
        Case "results"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN005")
        Case "winner"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN006")
        Case "bets"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN007")
        Case "raceoptions"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN001")
        Case "exceloptions"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN002")
        Case "language"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "LANGUAGE001")
        Case "info"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN009")
        Case "warning"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN010")
        Case "movie2017"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN011")
        Case "colours"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN021")
        Case "developer"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN022")
        Case "saverace"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN023")
        Case "loadrace"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN024")
        Case "replay"
            g_wksRace.OLEObjects(name).Object.caption = GetText(g_arr_Text, "BTN025")
    End Select
End Sub

'AI mode (Callback for Excel ribbon): Count the number of installed races
Private Sub AI_InstalledRaces_getItemCount(control As IRibbonControl, ByRef returnedVal)
    
    Dim wksR As Worksheet
    Dim cnt As Long
    
    Set g_colRacesInstalled = Nothing
    Set g_colRacesInstalled = New Collection
    
    For Each wksR In ThisWorkbook.Worksheets
        If left(wksR.name, 5) = "race_" Then
            If wksR.Cells(basAuxiliary.GetRow(wksR, "STATUS"), basAuxiliary.GetColumn(wksR, "RACE DATA VALUE")).Value = "released" Then cnt = cnt + 1
        End If
    Next wksR
    
    returnedVal = cnt
    
End Sub

'AI mode (Callback for Excel ribbon): Get the names of all installed races
Private Sub AI_InstalledRaces_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)

    Dim wksCheck As Worksheet
    Dim cnt As Long
    
    For Each wksCheck In ThisWorkbook.Worksheets
        If left(wksCheck.name, 5) = "race_" Then
            If wksCheck.Cells(basAuxiliary.GetRow(wksCheck, "STATUS"), basAuxiliary.GetColumn(wksCheck, "RACE DATA VALUE")).Value = "released" Then cnt = cnt + 1
            If cnt = index + 1 Then
                With wksCheck
                    g_colRacesInstalled.Add .name
                    returnedVal = .Cells(basAuxiliary.GetRow(wksCheck, "RACE NAME"), 2).Value & " " & _
                                    .Cells(basAuxiliary.GetRow(wksCheck, "YEAR"), 2).Value & " (" & _
                                    .Cells(basAuxiliary.GetRow(wksCheck, "DISTANCE METRES"), 2).Value & "m) - " & .Cells(basAuxiliary.GetRow(wksCheck, "TRACK LOCATION"), 2).Value
                    End With
                Exit For
            End If
        End If
    Next wksCheck

End Sub

'AI mode (Callback for Excel ribbon): Set default race
Private Sub AI_InstalledRaces_GetSelectedItemID(control As IRibbonControl, ByRef itemID As Variant)
    If objRace.SELECTED = "" Then
        objRace.SELECTED = g_colRacesInstalled(1) 'Take the first race of the collection
    End If
    itemID = objRace.SELECTED
End Sub

'AI mode (Callback for Excel ribbon): Set selected race
Private Sub AI_InstalledRaces_Click(control As IRibbonControl, id As String, index As Integer)
    objRace.SELECTED = g_colRacesInstalled(index + 1)
End Sub

'This procedure is executed when the workbook is being closed.
'The Auto_Close procedure can be used alternatively to the Workbook_BeforeClose
'event in "ThisWorkbook" ("DieseArbeitsmappe") which is NOT used for this project.
'If both procedures are implemented first the Workbook_BeforeClose is executed
'followed by Auto_Close.
Public Sub Auto_Close()
    'Reset Excel options
    Call basMainCode.ResetExcelOptions
    Application.ScreenUpdating = True
    If g_strPlayMode = "RS" Then
        Application.ReferenceStyle = objBasicData.XL_STYLE_NOTATION 'Restore the setting
        ThisWorkbook.Saved = True 'Do not save the workbook!
        'https://support.microsoft.com/en-us/help/213428/how-to-suppress-save-changes-prompt-when-you-close-a-workbook-in-excel
    End If
End Sub
