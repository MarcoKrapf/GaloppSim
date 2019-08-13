Attribute VB_Name = "basMainCode"
Option Explicit 'Force variable declaration
Option Private Module 'Prevent the public procedures in this module from being accessed from outside this module

'This module contains the main code of the horse racing simulator with most of the logic
'   Module basMainCode


'GALOPPSIM - Version 150.00 (August 2019)
'Horse Racing Simulator for Microsoft Excel
'Authors: Marco Matjes and Duncan Bölting
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
Public Const g_c_version As String = "(v150.00)" 'Version of the tool
Public Const g_c_email As String = "info@galoppsim.racing" 'Contact e-mail address
Public Const g_c_defaultRaceOptionsFile As String = "RACEOPTIONS" 'File name for race options
Public Const g_c_defaultFileType As String = ".GALOPPSIM" 'File type for GaloppSim files
Public Const g_c_errorLogFileName As String = "GALOPPSIM_ERRORLOG" 'File name for error logging
Public g_defaultPath As String 'Path for the GaloppSim files
Public g_strPlayMode As String ' "AI" = AddIn (.xlam) / "RS" = Run Simple (.xlsm)
Public g_RibbonGaloppSim As IRibbonUI 'Custom ribbon (only used for the AI edition)
Public AI_started As Boolean 'Will be set true in AI edition when activating the GALOPPSIM menu tab the first time
Public g_skipDelay As Boolean 'For skipping delay commands (Application.Wait)
Public g_errorLogPath As String 'Path for error logging
Public g_errorLogging As Boolean 'Error logging on/off
Public objRace As clsRace 'Object for all the data of a race
Public objOption As clsOptions 'Object for the race and excel options
Public objSpeed As clsSpeed 'Object for speed data
Public g_colRSbuttons As Collection 'Menu buttons in the RS edition
Public g_arr_varHorses() As Variant 'All information about the horses
Public g_colRacesInstalled As Collection 'List of installed races
Public g_oleComboRaces As OLEObject 'ComboBox with installed races in the RS edition
Public g_colBetSlips As Collection 'List with all betting slips
Public g_arr_Text() As Variant 'All text components
Public g_arr_Grammar(1 To 8) As String 'All animal grammar components
Public g_objLabel As MSForms.Label 'Label used for different purposes
Public g_strInpBoxReturnValue As String 'Return value of an input box
Public g_enumButton As String 'Return value of the pressed button

'Existing Worksheets with data
Public g_wksTEXT As Worksheet 'Worksheet with text components
Public g_wksPIC As Worksheet 'Worksheet with picture data
Public g_wksTEC As Worksheet 'Worksheet with technical data (speed, tactics...)

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


'VARIABLES AND CONSTANTS ON MODULE-LEVEL
'---------------------------------------
Dim m_wksCheck As Worksheet 'Variable to check whether the worksheet GALOPPSIM already exists
Dim m_arr_varPhotofinish() As Variant 'Position of each horse on the photo of the finish
Dim m_arr_varResultsCalc() As Variant 'Calculation of the position at the finish line
Dim m_arr_varResults() As Variant 'Race results list
Dim m_arr_varAdv() As Variant 'Advertising sequence
Dim m_intAdvertisingHeight As Integer 'Row height of the advertising area
Dim m_dblRankingsWidth As Double 'Cell width of the photo of the finish and the ranking list
Dim m_intTopRows As Integer 'Number of rows at the top of the worksheet (used for the menu in the RS edition)
Dim m_intLeftColumns As Integer 'Number of columns left of the starting gates
Dim m_intColumsAfterFinish As Integer 'Number of columns behind the finish line
Dim m_intTrackCellHeight As Integer 'Cell height (race track)
Dim m_dblTrackCellWidth As Double 'Cell width (race track)
Dim m_intFontSize As Integer 'Font size of the horse names and the hoof prints
Dim m_intHorsesRunning As Integer 'Number of horses currently running
Dim m_byteFavourite(1 To 3) As Byte 'Array for three predicted favourites of the race
Dim m_dblFavCalc(1 To 3) As Double 'Array for calculating the favourites
Dim m_intHorsesFinishing As Integer 'Variable that counts how many horses arrive at the finish line in one loop
Dim m_intFinishLoop As Integer 'Variable that counts in which computation period a placement was calculated
Dim m_strNameLeader As String 'Name of the leading horse
Dim m_intPosLeader As Integer 'Position of the leading horse
Dim m_strWinner As String 'Name(s) of the horse(s) at 1st place
Dim m_blnWin As Boolean 'Flag indicating that a horse has won the race
Dim m_blnPhotofinish As Boolean 'Flag indicating whether there is a photo finish
Dim m_blnDeadHeat As Boolean 'Dead heat (more than one horse are at 1st place)
Dim m_intPlace As Integer 'Placement in the finish
Dim m_arr_lngLugworms() As Long 'Array for the lugworm characters in a mudflats race
Dim i As Integer, j As Integer, k As Integer, m As Integer 'Counting variables for loops
Dim z As Long 'Auxiliary variable for loops


'PROCEDURES
'----------

'Main procedure for starting a new race
Private Sub NewRace()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    'Close pop-ups if visible
    If frmBettingAnalysis.Visible Then Unload frmBettingAnalysis 'Analysis of the bets
    If frmRS_navigation.Visible Then Unload frmRS_navigation 'Navigation panel (RS edition only)
    
'    If g_strPlayMode = "AI" Then 'Kann das RAUS ? TEST in 150.00
'        Call GetTextComponents 'Get texts according to the selected language
'    End If
    
    'Reset the betting slip collection
    Set g_colBetSlips = Nothing
    Set g_colBetSlips = New Collection

    Call GetRaceData 'Get the race data from the worksheet with the selected race
    Call GetAnimalGrammar 'Get grammar components according to the selected language
    Call AssignBasicValues 'Get basic values
    Call GetHorseData 'Get data about the horses according to the selected race
    Call CheckBettingAllowed 'Check whether betting is allowed for this race
    Call ShowStartPopup 'Show the pop-up with the race to run
        
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
                    .Value = g_c_tool & " " & g_c_version
                End With
            Call RS_MenuAreaHide 'Hide the controls on the worksheet
        End If

        Call DrawRaceTrack 'Race track generation
        Call DrawHorseNames 'Write the horse names on the race track if selected
        If objOption.MOMENTUM Then Call MomentumFormattings 'Prepare the race sheet for speed bars
        If objOption.RACE_INFO And objOption.RACE_INFO_POP Then Call RaceInfoPopup 'Pop-up with race the info if selected
        Call CheckSpaceRace 'Check whether a space race is chosen
        Call RaceWelcome 'Pop-up with a warm welcome to the race
        Call StartingGrid 'Put the horses in the gates
        Call RacePresentation 'Presentation of the horses
        Call RunRace 'Race start
        Call RankNotFinished 'Find the horses that did not finish
        Call RaceFinished 'Info pop-up when the race is over
        Call CheckDeadHeat 'Check whether more than one horse has won
        Call DrawRankingList(True) 'Race results on the ranking list
        Call DrawWinnerPhoto 'Show a photo of the winner
        If objOption.BET_PLACED And objOption.BET_ANALYSIS Then Call AnalyseBettings 'Pop-up with the bet slips analysis
        
        If g_strPlayMode = "RS" Then 'Show the navigation panel (RS edition only)
            Call basAuxiliary.ActivateRaceSheet
            'Activate buttons
            With g_wksRace
                .OLEObjects("finishphoto").Object.Enabled = True
                .OLEObjects("results").Object.Enabled = True
                .OLEObjects("winner").Object.Enabled = True
                If objOption.BET_PLACED Then
                    .OLEObjects("bets").Object.Enabled = True
                End If
                .rows(1).RowHeight = 8
                .rows(2).RowHeight = 18
            End With
            Call RS_MenuAreaShow(True)
        End If

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

'Starting procedure (called by the Workbook_Open event in the RS edition)
Public Sub RS_NewRace()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    'As this procedure is triggered automatically and executed only once
    'when opening the workbook comment in the "Stop" command in the
    'following line for debugging purposes
'    Stop
    
    m_intTopRows = 8 'Space on the top of the worksheet for the control elements
                        'like the dropdown and menu buttons
    Call CreateRaceSheet 'Create the worksheet "GALOPPSIM"
    Call RS_StartScreen 'Draw the start screen
    Call RS_AddControls 'Add menu buttons and a dropdown for the race selection

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RS_NewRace()")
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
Private Sub PaintHorse(ByVal row As Integer, tail As Integer, color As Variant)
    Dim horseColor As Long
    Dim check_array As Boolean
    
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Call basAuxiliary.ActivateRaceSheet 'Ensure that the "GALOPPSIM" worksheet is activated
    
    check_array = False 'Reset the check variable
    If IsArray(color) Then 'Multicoloured horse
        If UBound(color) > 6 Then
            Dim count As Integer, countInitial As Integer
            Dim currentColor As Long
            count = 0 'Reset the segment counter
            Do While count < 8 'Loop through the 8 segments of a horse
                countInitial = count 'Set the counter to the current segment
                currentColor = color(count) 'Colour code of the current segment
                
                Do 'Check whether adjoining segments have the same colour
                    If count = 7 Then Exit Do
                    If currentColor = color(count + 1) Then count = count + 1 Else Exit Do
                Loop
                If count = countInitial Then 'The current segment has a different colour as the next
                    Cells(row, tail + count).Interior.color = currentColor 'Paint a single horse segment
                Else 'Some consecutive segments have the colour and can be painted as a block for better performance
                    Range(Cells(row, tail + countInitial), Cells(row, tail + count)).Interior.color = currentColor
                End If
                count = count + 1 'Next segment
            Loop
            check_array = True
        End If
    End If
    
    If Not check_array Then 'Monochrome horse (no array submitted)
        If IsNumeric(color) Then 'A number with the colour code has been provided
            horseColor = color
        Else 'No valid color has been provided, use brown instead
            horseColor = 3291720
        End If
            Range(Cells(row, tail), Cells(row, tail + 7)).Interior.color = horseColor 'Paint the horse as a whole
    End If
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "PaintHorse()")
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
Private Sub CheckSpaceRace()
    'Check whether the Race ID begins with "SPACE"
    If objRace.RACE_ID <> "SPACE" Then Exit Sub
    
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
    strSpaceInfo = GetText(g_arr_Text, "SPACEINFO001") & vbNewLine
    'No slipstream in space
    strSpaceInfo = strSpaceInfo & GetText(g_arr_Text, "SPACEINFO002") & vbNewLine
    'No start refusal in space
    strSpaceInfo = strSpaceInfo & GetText(g_arr_Text, "SPACEINFO003")
    
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
            background-color (hex), caption with the initial text
        Call RS_AddButton("raceoptions", 15, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN001"))
        Call RS_AddButton("exceloptions", 99, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN002"))
        Call RS_AddButton("startrace", 196, 40, 81, 49, 11, True, 52377, GetText(g_arr_Text, captionStart)) 'Green button
        Call RS_AddButton("finishphoto", 280, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN004"))
        Call RS_AddButton("results", 364, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN005"))
        Call RS_AddButton("winner", 448, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN006"))
        Call RS_AddButton("bets", 532, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN007"))
        Call RS_AddButton("language", 629, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "LANGUAGE001"))
        Call RS_AddButton("info", 713, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN009"))
        Call RS_AddButton("warning", 797, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN010"))
        Call RS_AddButton("movie2017", 881, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN011"))
        
    'Add a combobox for the race selection to the menu area
        Call RS_AddComboboxRaces("CBraces", 196, 15, 597, 22) '"name(ID)", left, top, width, height
    
    'Deactivate some buttons as they have no function before the race
        Call RS_InactivateCommandButtons

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RS_AddControls()")
    Call basAuxiliary.CodeCrash
End Sub

'Set some buttons inactive as they have no function before the race
Private Sub RS_InactivateCommandButtons()
    Call basAuxiliary.ActivateRaceSheet
        With g_wksRace
            .OLEObjects("finishphoto").Object.Enabled = False
            .OLEObjects("results").Object.Enabled = False
            .OLEObjects("winner").Object.Enabled = False
            .OLEObjects("bets").Object.Enabled = False
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
                
                Call ShowNewRaceScreen 'Show the GALOPPSIM title screen
            End If
            
            Call GetRace 'Get the selected race from the dropdown menu
            Call GetRaceData 'Read the data of the selected race from the worksheet
            Call RS_InactivateCommandButtons 'Inactivate some buttons (RS edition only)
            Call NewRace 'Main procedure for starting a new race
        
        Case "finishphoto"
            Call ShowFinishPhoto
        Case "results"
            Call ShowRankingList
        Case "winner"
            Call ShowWinnerPhoto
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
                Call GetRaceData 'Read the data of the selected race from the worksheet
                Call GetAnimalGrammar 'Get grammar components
            End If
            Call ChangeLanguage
        Case "info"
            Call ShowInfo
        Case "warning"
            Call ShowWarning
        Case "movie2017"
            Call GaloppSimMovie2017
    End Select
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RS_ExecuteClick()")
    Call basAuxiliary.CodeCrash
End Sub

'Show the GALOPPSIM title screen
Private Sub ShowNewRaceScreen()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Call basAuxiliary.ActivateRaceSheet
    
    'Adjust the column width and row height according to the screen size
    With g_wksRace.UsedRange
        .ColumnWidth = ZoomLevelPictures()(0)
        .RowHeight = ZoomLevelPictures()(1)
        .Clear
    End With
    
    If g_strPlayMode = "RS" Then Call RS_MenuAreaHide 'Hide the controls on the worksheet (RS edition)
    If g_strPlayMode = "AI" Then Call AI_ExcelModeStart
    ActiveWindow.ScrollColumn = 1 'Scroll to the left (column A)
    Call PaintPicture(g_wksPIC, g_wksRace, "NEWRACE", 100, 40, 1, 1) 'Paint the GALOPPSIM title picture
    Call CursorAway 'Place the cursor far away (in the upper right corner of the screen)
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "ShowNewRaceScreen()")
    Call basAuxiliary.CodeCrash
End Sub

'Add a new button to the worksheet in the RS edition
Private Sub RS_AddButton(n As String, l As Integer, t As Integer, w As Integer, _
                            h As Integer, fs As Integer, fb As Boolean, bc As Long, c As String)
            'n = name(ID), l = left, t = top, w = width, h = height, fs = font-size,
            'fb = font:bold, bc = background-color (hex), c =caption with the initial text
    
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
        .Object.ColumnCount = 2 'column 0: race name
        .Object.ColumnWidths = "0 Pt" 'width of the column with the race name --> hidden
        .Object.Style = fmStyleDropDownList 'Allow only values from the item list, no free entries
        .Visible = True
    End With

    'Populate the dropdown with all installed races (all worksheets beginning with "race_")
    For Each wksCheck In ThisWorkbook.Worksheets
        If left(wksCheck.name, 5) = "race_" Then
            If wksCheck.Cells(basAuxiliary.GetRow(wksCheck, "STATUS"), 2).Value = "released" Then
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
    Range(rows(1), rows(m_intTopRows)).Hidden = True
End Sub

'Show all controls on the worksheet in the RS edition
Public Sub RS_MenuAreaShow(popup As Boolean)
    Dim oleObj As OLEObject
    
    Call basAuxiliary.ActivateRaceSheet

    'Show the top rows which are used for the control objects
    Range(rows(1), rows(m_intTopRows)).Hidden = False

    'Loop through all control objects and show them one by one
    For Each oleObj In g_wksRace.OLEObjects
        #If Debugging Then
            Debug.Print "Show OLEObject: " & oleObj.name
        #End If
        oleObj.Visible = True
    Next oleObj

    If popup = True Then frmRS_navigation.show (vbModeless) 'Show the pop-up with the navigation panel

End Sub

'Start screen for the RS edition
Private Sub RS_StartScreen()

    'As this procedure is triggered automatically and executed only once
    'when opening the workbook comment in the "Stop" command in the
    'following line for debugging purposes
'    Stop
    
    'Paint the title picture
    Call PaintPicture(g_wksPIC, g_wksRace, "RUNSIMPLE", 100, 40, 1, 1)
    'Formattings for the area with the picture
    With Range(Columns(1), Columns(100)) 'Column width and row height dependent on the window size
        .ColumnWidth = ZoomLevelPictures()(0)
        .RowHeight = ZoomLevelPictures()(1)
        .rows(1).RowHeight = 8 'Height of the top row
        With .rows(2) 'Formattings for the second row
            .Font.name = "Arial Black" 'For getting the font in the dropdown with the races bold
            .EntireRow.AutoFit
        End With
    End With
    'Write the title in cell B2
    Cells(2, 2).Value = g_c_tool & " " & g_c_version
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
Private Sub AssignBasicValues()
    On Error GoTo ERRORHANDLING 'In case an error occurs
        
    'Determine the zoom level
    If objOption.ZOOM_LEVEL = 0 Then 'If there is no value set yet
        objOption.ZOOM_LEVEL = ZoomLevelRecommendation() 'Take the recommended value
    Else
        Dim byteOpt As Byte
        byteOpt = ZoomLevelRecommendation() 'Get a recommendation

        If objOption.ZOOM_LEVEL <> byteOpt Then 'Compare the selected value with the recommendation
            Call ShowMessagePopup(objRace.RACE_NAME & " " & objRace.RACE_YEAR, _
                GetText(g_arr_Text, "ZOOM007") & vbNewLine & vbNewLine & _
                GetText(g_arr_Text, "ZOOM008") & ": " & ZoomLevelText(objOption.ZOOM_LEVEL) & vbNewLine & _
                GetText(g_arr_Text, "ZOOM009") & ": " & ZoomLevelText(byteOpt) & vbNewLine & vbNewLine & _
                GetText(g_arr_Text, "ZOOM010"), _
                enumButton.YesNo, vbModal)
                    
                'Evaluate the return value of the clicked button
                If g_enumButton = enumButton.yes Then objOption.ZOOM_LEVEL = byteOpt 'Adapt the value
        End If
    End If
    
    'Assign the values according to the zoom level
        Select Case objOption.ZOOM_LEVEL
            Case 1
                m_intTrackCellHeight = 6 'Cell height on the race track
                m_dblTrackCellWidth = 0.2 'Cell width (i.e. length) on the race track
                m_intFontSize = 5 'Font size for the horse names and the hoof prints
                m_dblRankingsWidth = 3 'Cell width for the finish photo and the ranking list
                m_intAdvertisingHeight = 2 'Row height in the advertising area
            Case 2
                m_intTrackCellHeight = 9
                m_dblTrackCellWidth = 0.3
                m_intFontSize = 8
                m_dblRankingsWidth = 4.5
                m_intAdvertisingHeight = 5
            Case 3
                m_intTrackCellHeight = 12
                m_dblTrackCellWidth = 0.4
                m_intFontSize = 10
                m_dblRankingsWidth = 6
                m_intAdvertisingHeight = 6
        End Select
        
'Version 150.10: At runtime auslesen
    If objRace.RACE_ID = "SPACE" Then
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
    m_intLeftColumns = 11
    'Columns behind the finish line (minimum 5)
    m_intColumsAfterFinish = 5
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "AssignBasicValues()")
    Call basAuxiliary.CodeCrash
End Sub

'Read the data of the selected race from the worksheet
Private Sub GetRaceData()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    'Assign the worksheet with the selected race
    Set g_wksRaceData = ThisWorkbook.Worksheets(objRace.SELECTED)
    
    'Read the full race data from the worksheet
    With g_wksRaceData
        objRace.RACE_ID = .Cells(basAuxiliary.GetRow(g_wksRaceData, "RACE ID"), 2).Value 'Unique race ID
        objRace.REAL_RACE = .Cells(basAuxiliary.GetRow(g_wksRaceData, "REAL RACE"), 2).Value 'Real race (yes or no)
        objRace.PARTICIPANTS = .Cells(basAuxiliary.GetRow(g_wksRaceData, "PARTICIPANTS"), 2).Value 'Race participants (HORSE/PIG/DONKEY/DOG/UNICORN)
        objRace.RACE_NAME = .Cells(basAuxiliary.GetRow(g_wksRaceData, "RACE NAME"), 2).Value 'Race name
        objRace.RACE_YEAR = .Cells(basAuxiliary.GetRow(g_wksRaceData, "YEAR"), 2).Value 'Year of the race
        objRace.TRACK_LOCATION = .Cells(basAuxiliary.GetRow(g_wksRaceData, "TRACK LOCATION"), 2).Value 'Track location
        objRace.COUNTRY = GetCountryName(.Cells(basAuxiliary.GetRow(g_wksRaceData, "COUNTRY"), 2), objOption.language) 'Country
        objRace.TRACK_NAME = .Cells(basAuxiliary.GetRow(g_wksRaceData, "TRACK NAME"), 2).Value 'Track name
        objRace.TRACK_COLOUR = .Cells(basAuxiliary.GetRow(g_wksRaceData, "TRACK COLOUR"), 2).Value 'Track colour
        objRace.TRACK_SURFACE = .Cells(basAuxiliary.GetRow(g_wksRaceData, "TRACK SURFACE"), 2).Value 'Track surface
        objRace.RACE_TYPE = .Cells(basAuxiliary.GetRow(g_wksRaceData, "RACE TYPE"), 2).Value 'Race type
        objRace.METRES = .Cells(basAuxiliary.GetRow(g_wksRaceData, "DISTANCE METRES"), 2).Value 'Race distance
        objRace.STARTING_GATE = .Cells(basAuxiliary.GetRow(g_wksRaceData, "STARTING GATE"), 2).Value 'Starting gate (yes or no)
        objRace.LANES_FIX_OR_RANDOM = .Cells(basAuxiliary.GetRow(g_wksRaceData, "LANES FIX OR RANDOM"), 2).Value 'Lanes fix or random
        objRace.ADVERTISING = .Cells(basAuxiliary.GetRow(g_wksRaceData, "ADVERTISING"), 2).Value 'Advertising (yes or no)
        objRace.BETTING_ALLOWED = .Cells(basAuxiliary.GetRow(g_wksRaceData, "BETTING ALLOWED"), 2).Value 'Betting allowed (yes or no)
        objRace.NUMBER_ENROLLED = .Cells(.rows.count, GetColumn(g_wksRaceData, "STATUS")).End(xlUp).row - 1 'Number of horses enrolled
        objRace.NUMBER_STARTING = Application.WorksheetFunction.CountIf(.Columns(GetColumn(g_wksRaceData, "STATUS")), "START")
    End With
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "GetRaceData()")
    Call basAuxiliary.CodeCrash
End Sub

'Get the data for advertisement in a race
Private Sub GetAdvertisementData()
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Dim col As Integer 'Column with the advertisement data
    col = basAuxiliary.GetColumn(g_wksRaceData, "ADVERTISEMENT")

        j = g_wksRaceData.Cells(rows.count, col).End(xlUp).row - 1 'Last row with data in the column ADVERTISEMENT
        ReDim m_arr_varAdv(1 To j) 'Location of the advertisement data
        For i = 1 To j
            For k = 1 To g_wksPIC.Cells(1, Columns.count).End(xlToLeft).Column 'Loop through the columns with picture data
                If g_wksRaceData.Cells(i + 1, col).Value = g_wksPIC.Cells(1, k).Value Then 'Find the column on the worksheet "PIC"
                    m_arr_varAdv(i) = k 'Write the column number into an array
                    Exit For
                End If
            Next k
        Next i
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "GetAdvertisementData()")
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
        GetCountryName = g_wksTEXT.Cells(row, 5).Value 'Take the English name as fallback
    Else
        GetCountryName = g_wksTEXT.Cells(row, col).Value 'Get the name according to the selected language
    End If
End Function

'Read all horse data from the race sheet
Private Sub GetHorseData()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    'Resize the arrays for the horse data
    ReDim g_arr_varHorses(1 To objRace.NUMBER_ENROLLED, 0 To 31) 'All data of the horses
    ReDim m_arr_varPhotofinish(1 To objRace.NUMBER_ENROLLED, 0 To 4) 'Snapshot of the finish
    ReDim m_arr_varResults(0 To objRace.NUMBER_STARTING, 0 To 7) 'Ranking list
    
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
        Dim arr_color(0 To 7) 'Array with the colours of the 8 horse segments
        Dim color As Integer
        Dim same_color As Boolean
        same_color = True
        g_arr_varHorses(i, 0) = g_wksRaceData.Cells(1 + i, 6).Value 'Status (START, CANCELLED, REFUSED...)
        g_arr_varHorses(i, 11) = g_wksRaceData.Cells(1 + i, 5).Value 'Horse number
        g_arr_varHorses(i, 1) = g_wksRaceData.Cells(1 + i, 7).Value 'Horse name
        
        'Get the horse colours
        If Not IsEmpty(g_wksRaceData.Cells(1 + i, 15)) Then 'If the colour of the head is not empty
            For color = 0 To 7
                arr_color(color) = g_wksRaceData.Cells(1 + i, 8 + color) 'Segment colour
                If color > 0 Then 'If the colour is not empty (or black)
                    If Not g_wksRaceData.Cells(1 + i, 8 + color) = g_wksRaceData.Cells(1 + i, 7 + color) Then
                        same_color = False 'If the colour of this segment differs from that before
                    End If
                End If
            Next color
            'Assign either a Long value or an Array with 8 fields for the horse colour
            If same_color Then g_arr_varHorses(i, 2) = arr_color(0) Else g_arr_varHorses(i, 2) = arr_color
        Else 'If no colour of the head is found: Determine a random colour for the whole horse
            Randomize 'Initialize the random number generator
            g_arr_varHorses(i, 2) = Int((16777215 - 0 + 1) * Rnd + 0) 'Apply the randomly generated colour for the whole horse
            'Formula for the generation of a random integer value within a specific range:
            'Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
            '>> replace upperbound and lowerbound with integer values
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
            g_arr_varHorses(i, 15) = g_wksRaceData.Cells(1 + i, 4).Value 'Read the starting gate from the worksheet
        End If
        
        g_arr_varHorses(i, 3) = m_intTopRows + 5 + 2 * g_arr_varHorses(i, 15) 'Row number on which the horse will run
        
        g_arr_varHorses(i, 4) = m_intLeftColumns + 12 'Starting position (column number)
        g_arr_varHorses(i, 9) = 0 'Exact position for internal calculation (0 to [race distance in metres * 100])
        
        'Get the basic speed
        If Not IsEmpty(g_wksRaceData.Cells(1 + i, 16)) Then 'If a value is found on the race sheet
            g_arr_varHorses(i, 5) = g_wksRaceData.Cells(1 + i, 16).Value 'Fixed basic speed
        Else 'If no value is found: generate it by random within a determined range
            Randomize
            g_arr_varHorses(i, 5) = Int((objSpeed.SPEED_BASIC_HIGH - objSpeed.SPEED_BASIC_LOW + 1) * Rnd + objSpeed.SPEED_BASIC_LOW)
        End If
        
        'Determine the daily form of the horse by random
            Randomize
            g_arr_varHorses(i, 6) = Int((objSpeed.SPEED_COND_HIGH - objSpeed.SPEED_COND_LOW + 1) * Rnd + objSpeed.SPEED_COND_LOW)
        
        'Determine the betting odds
            If Not IsEmpty(g_wksRaceData.Cells(1 + i, 17)) Then 'If a value is found on the race sheet
                g_arr_varHorses(i, 17) = g_wksRaceData.Cells(1 + i, 17).Value 'Fixed odds
            Else 'If no value is found: derive it from the basic speed with a complex formula
                Randomize
                'Rounded integer value from (((50 + (((number of starters + 2) / 6) * (1523 - basic speed) ^ 2)) / 5) * random value between 0.9 and 1.1)
                g_arr_varHorses(i, 17) = Round(((50 + (((objRace.NUMBER_ENROLLED + 2) / 6) * (1523 - g_arr_varHorses(i, 5)) ^ 2)) / 5) * (Int((11 - 9 + 1) * Rnd + 9) / 10), 0)
            End If
            
        'Estimation error for the impression during the warm-up (+/-50 pixels of the bar length)
            Randomize
            g_arr_varHorses(i, 18) = (Int((100 - 0 + 1) * Rnd + 0)) - 50 'Random number between -50 and +50
            
        'Reset the slipstream factor
            g_arr_varHorses(i, 22) = 0
            
        'Get the picture of the winner
        If g_wksRaceData.Cells(1 + i, 18).Value <> "" Then 'If a value is found on the race sheet
            g_arr_varHorses(i, 23) = g_wksRaceData.Cells(1 + i, 18).Value 'Specific picture
        Else 'If no value is found: take the default picture
            g_arr_varHorses(i, 23) = "WINNER_" & objRace.PARTICIPANTS & "_DEFAULT"
        End If
        
        'Attribute for different purposes like a special race behaviour
            g_arr_varHorses(i, 24) = g_wksRaceData.Cells(1 + i, 19).Value
    Next i
    
    'Calculation of the race favourites
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
                
    'If race tactics are activated: Determine the speed of each horse in each phase of the race
    If objOption.TACTICS = True Then
    
        Dim a As Integer
        Dim aux() As String
        
        For i = 1 To objRace.NUMBER_ENROLLED
                
            If g_wksRaceData.Cells(1 + i, 27).Value <> "" Then 'If a value is found on the race sheet
                g_arr_varHorses(i, 25) = g_wksRaceData.Cells(1 + i, 27).Value
            Else
                g_arr_varHorses(i, 25) = "SSMMFF"
                ReDim aux(1 To 6)
                For j = 1 To 6
                    Do
                        a = Int((6 - 1 + 1) * Rnd + 1) 'Platz im Array finden
                        If aux(a) = "" Then
                            aux(a) = Mid(g_arr_varHorses(i, 25), j, 1)
                            Exit Do
                        End If
                    Loop
                Next
                g_arr_varHorses(i, 25) = ""
                For j = 1 To 6
                    g_arr_varHorses(i, 25) = g_arr_varHorses(i, 25) & aux(j)
                Next
            End If
            
            For j = 1 To 6
                g_arr_varHorses(i, 25 + j) = TacticMapping(Mid(g_arr_varHorses(i, 25), j, 1))
            Next j
        
        Next i
    End If
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "GetHorseData()")
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
'        If g_strPlayMode = "RS" Then 'ToDo... Muss das rein für AI ? TEST in 150.00
            .DisplayFullScreen = blnFull 'Excel ribbon
'        End If
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
Public Function ZoomLevelPictures() As Variant() 'return an array with the values for the column width and row height
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

'Zoom level recommendation dependent on the window size and the number of starters
Public Function ZoomLevelRecommendation() As Byte

    Dim dblWindowHeight As Double
    Dim intHorses As Byte

    dblWindowHeight = Application.ActiveWindow.Height 'Window height
    
    'Get the number of horses in the selected race
    intHorses = ThisWorkbook.Worksheets(objRace.SELECTED).Cells( _
        basAuxiliary.GetRow(ThisWorkbook.Worksheets(objRace.SELECTED), "NUMBER OF STARTERS"), 2).Value

    Select Case dblWindowHeight
        Case Is > 790 'Window height (e.g. screen resolution 1680x1050 pixels)
            If intHorses <= 23 Then 'Up to 23 horses
                ZoomLevelRecommendation = 3
            ElseIf intHorses <= 31 Then '24-31 horses
                ZoomLevelRecommendation = 2
            Else 'More than 31 horses
                ZoomLevelRecommendation = 1
            End If
        Case Is > 550 'Window height (e.g. screen resolution 1368x768 pixels)
            If intHorses <= 9 Then 'Up to 9 horses
                ZoomLevelRecommendation = 3
            ElseIf intHorses <= 14 Then '10-14 horses
                ZoomLevelRecommendation = 2
            Else 'More than 14 horses
                ZoomLevelRecommendation = 1
            End If
        Case Else 'Small
            ZoomLevelRecommendation = 1
    End Select
End Function

'Retrieve the zoom level text
Public Function ZoomLevelText(byteZL As Byte) As String
    Select Case byteZL
        Case 1
            ZoomLevelText = GetText(g_arr_Text, "ZOOM003") 'Small
        Case 2
            ZoomLevelText = GetText(g_arr_Text, "ZOOM004") 'Medium
        Case 3
            ZoomLevelText = GetText(g_arr_Text, "ZOOM005") 'Large
    End Select
End Function

'Retrieve data for the horse size preview
Public Function HorseSizePreview(intHSP) As Variant() 'Return an array with the values
    Select Case intHSP
        Case 1
            HorseSizePreview = Array(12, 6) '(Width, height) of the horse (small)
        Case 2
            HorseSizePreview = Array(22, 9) '(medium)
        Case 3
            HorseSizePreview = Array(30, 12) '(large)
    End Select
End Function

'Grammar components for compiling texts with different race participants
Public Sub GetAnimalGrammar()
    Dim animal As String 'Race participants (HORSE/PIG/DONKEY/DOG/UNICORN)
    Dim col As Integer

    With ThisWorkbook 'Read the type of participants from the worksheet
        animal = .Worksheets(objRace.SELECTED).Cells(basAuxiliary.GetRow(.Worksheets(objRace.SELECTED), "PARTICIPANTS"), 2).Value
    End With
    
    'Get the column with the language
    col = basAuxiliary.GetColumn(g_wksTEXT, objOption.language)
    
    'Get the text component
    For i = 1 To g_wksTEXT.Cells(rows.count, 1).End(xlUp).row 'ToDo... geht das in AI? Ansonsten hart auf 1000 setzen TEST in 150.00
        If g_wksTEXT.Cells(i, 1).Value = animal Then
            g_arr_Grammar(g_wksTEXT.Cells(i, 2).Value) = g_wksTEXT.Cells(i, col).Value
        End If
    Next i
    
End Sub

'Draw the race track when starting a new race
Private Sub DrawRaceTrack()
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Call basAuxiliary.ActivateRaceSheet
    
    'Deactivate screen updating
    Application.ScreenUpdating = False
        
    'Freeze columns A-L if one of those checkboxes is activated, otherwise unfreeze
    If objOption.NAMES_LEFT Or objOption.COLOURS_LEFT Or objOption.HIGHLIGHT_FAV _
        Or (objOption.FOCUSED_RUN And objOption.HIGHLIGHT_FOC) Or (objOption.RACE_INFO And objOption.RACE_INFO_WKS) Then
            Call basAuxiliary.Freeze(12, 0, True) 'Freeze
    Else
            Call basAuxiliary.Freeze(0, 0, False) 'Unfreeze
    End If

    'Formatting: Row height of different sections
    Range(rows(1 + m_intTopRows), rows(5 + m_intTopRows)).EntireRow.RowHeight = 15 'Above the race track
    Range(rows(objRace.NUMBER_ENROLLED * 2 + 6 + 1 + m_intTopRows), _
        rows(objRace.NUMBER_ENROLLED * 2 + 52 + m_intTopRows)).EntireRow.RowHeight = 15 'Below the race track
    rows(objRace.NUMBER_ENROLLED * 2 + 20 + m_intTopRows).RowHeight = 20 'Headline of the ranking list

    'Display race data on the top
    With Cells(2 + m_intTopRows, 14) 'Race name, year, track and location
        .Font.name = "Arial Black"
        .Value = objRace.RACE_NAME & " " & objRace.RACE_YEAR & " - " & objRace.TRACK_NAME & ", " & objRace.TRACK_LOCATION _
            & " (" & objRace.COUNTRY & ")"
    End With
    With Cells(3 + m_intTopRows, 14) 'Race and track type
        .Font.name = "Arial"
        .Font.Bold = True
        .Value = objRace.RACE_TYPE_TEXT & " " & GetText(g_arr_Text, "RACE007") & " " & _
            objRace.METRES & " " & GetText(g_arr_Text, "RACE009") & " - " & objRace.TRACK_SURFACE_TEXT
    End With
    
    'Formatting: Columns on the left of the starting grid dependent on the zoom level
    Columns(1).ColumnWidth = 2 * objOption.ZOOM_LEVEL 'Left margin
    Range(Columns(2), Columns(9)).ColumnWidth = m_dblTrackCellWidth 'Horse colours
    Columns(10).ColumnWidth = objOption.ZOOM_LEVEL 'Empty column
    Columns(11).ColumnWidth = 2 * objOption.ZOOM_LEVEL 'Horse numbers
    Columns(12).ColumnWidth = 20 + (objOption.ZOOM_LEVEL * 6) 'Horse names
    Range(Columns(13), Columns(m_intLeftColumns + 12)).ColumnWidth = m_dblTrackCellWidth 'Starting area
    Columns(m_intLeftColumns + 4).ColumnWidth = 3 + objOption.ZOOM_LEVEL 'Starting gate numbers
        
    'Formatting: Race track to run (1 metre = 1 column)
    Range(Columns(m_intLeftColumns + 13), Columns(objRace.METRES + m_intLeftColumns + 13 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))).ColumnWidth = m_dblTrackCellWidth 'Column width
    For i = (6 + m_intTopRows) To (objRace.NUMBER_ENROLLED * 2 + 6 + m_intTopRows)
        rows(i).RowHeight = m_intTrackCellHeight / (2 - (i - (6 + m_intTopRows)) Mod 2) 'Row height: Alternating higher and lower
    Next i
    Range(Cells(4 + m_intTopRows, 1), Cells(objRace.NUMBER_ENROLLED * 2 + 19 + m_intTopRows, objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))).Interior.color = objRace.TRACK_COLOUR 'Track colour
    With Range(Cells(4 + m_intTopRows, 1), Cells(objRace.NUMBER_ENROLLED * 2 + 8 + m_intTopRows, objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))) 'Track font
        .Font.name = "Arial"
        .Font.size = m_intFontSize
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    'In case of using starting gates
    If objRace.STARTING_GATE = "Y" Then
        'Draw the gates
            For i = 6 To (objRace.NUMBER_ENROLLED * 2 + 6) Step 2 'One gate for each starting place
                Range(Cells(i + m_intTopRows, m_intLeftColumns + 8), Cells(i + m_intTopRows, m_intLeftColumns + 13)).Interior.ColorIndex = 1
            Next i
            Range(Cells(6 + m_intTopRows, m_intLeftColumns + 13), Cells(objRace.NUMBER_ENROLLED * 2 + 6 + m_intTopRows, m_intLeftColumns + 13)).Interior.ColorIndex = 1 'Close the gates
        'Label the gates
            With Range(Cells(7 + m_intTopRows, m_intLeftColumns + 4), Cells(objRace.NUMBER_ENROLLED * 2 + 5 + m_intTopRows, m_intLeftColumns + 4))
                .Font.ColorIndex = 1
                .Font.size = m_intFontSize
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
            End With
            For i = 1 To (objRace.NUMBER_ENROLLED) 'Gate numbers
                Cells(5 + 2 * i + m_intTopRows, m_intLeftColumns + 4).Value = GetText(g_arr_Text, "RACE010") & " " & i
            Next i
    End If
        
    'Display metres above and below the race track
    For i = objOption.METRES_DISPLAY To (objRace.METRES - 20) Step objOption.METRES_DISPLAY
        
        #If Debugging Then 'For debugging purposes: Vertical line at each marker position
            Range(Cells(5, i + m_intLeftColumns + 11), Cells(45, i + m_intLeftColumns + 11)).Interior.ColorIndex = 1
        #End If
        
        With Cells(4 + m_intTopRows, i + m_intLeftColumns + 11) 'Above the track
            .Font.name = "Arial"
            .Font.Bold = True
            .Value = i & GetText(g_arr_Text, "RACE008") '"m"
        End With
        With Cells(objRace.NUMBER_ENROLLED * 2 + 8 + m_intTopRows, i + m_intLeftColumns + 11) 'Below the track
            .Font.name = "Arial"
            .Font.Bold = True
            .Value = i & GetText(g_arr_Text, "RACE008") '"m"
        End With
    Next i
    
    'Formatting: horse names on the left
    With Range(Cells(6 + m_intTopRows, 11), Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, 12))
        .Font.color = objRace.TRACK_COLOUR 'Track colour, so that the names are not visible yet
        .IndentLevel = 1 'Text indented
        .Font.size = m_intFontSize
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
        
    'In case of a mudflats race
    If objRace.TRACK_SURFACE = "M" Then Call DrawMudflats
        
    'Formatting: finish area
    Columns(objRace.METRES + m_intLeftColumns + 12).ColumnWidth = m_dblTrackCellWidth 'Width of the finish line
    Range(Cells(5 + m_intTopRows, objRace.METRES + m_intLeftColumns + 12), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METRES + m_intLeftColumns + 12)).Interior.ColorIndex = 56 'Colour of the finish line: dark grey
    With Cells(4 + m_intTopRows, objRace.METRES + m_intLeftColumns + 11) 'Race distance above the track
        .Font.name = "Arial"
        .Font.Bold = True
        .Value = objRace.METRES & GetText(g_arr_Text, "RACE008")
    End With
    With Cells(objRace.NUMBER_ENROLLED * 2 + 8 + m_intTopRows, objRace.METRES + m_intLeftColumns + 11)  'Race distance below the track
        .Font.name = "Arial"
        .Font.Bold = True
        .Value = objRace.METRES & GetText(g_arr_Text, "RACE008")
    End With
        
    'Formatting: area behind the finish line
    Columns(objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).ColumnWidth = 18 + (objOption.ZOOM_LEVEL * 6)
    Columns(objRace.METRES + m_intLeftColumns + 15 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).ColumnWidth = objOption.ZOOM_LEVEL * 3
    Range(Columns(objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), Columns(objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR) + 300)).ColumnWidth = m_dblTrackCellWidth 'Column width
    
    'Formatting: horse names behind the finish line
    With Range(Cells(5 + m_intTopRows, objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), Cells(objRace.NUMBER_ENROLLED * 2 + 7 + (2 * objOption.SPEED_FACTOR) + m_intTopRows, objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)))
        .Font.ColorIndex = 1
        .IndentLevel = 1
        .Font.size = m_intFontSize
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    'Advertising area below the race track
    Range(rows(objRace.NUMBER_ENROLLED * 2 + 9 + m_intTopRows), _
        rows(objRace.NUMBER_ENROLLED * 2 + 19 + m_intTopRows)).EntireRow.RowHeight = m_intAdvertisingHeight 'Row height according to the zoom level

    If objRace.ADVERTISING = "Y" Then
        Dim advPos As Integer 'Column position for the next ad
        
        'Get advertisements from the race sheet
        Call GetAdvertisementData

        advPos = m_intLeftColumns + 12
        For i = 1 To UBound(m_arr_varAdv) 'Loop through the array which contains the ads
            z = 3 'Set the pointer to the first colour code
            For j = advPos To advPos + g_wksPIC.Cells(2, m_arr_varAdv(i)) - 1
                If j >= objRace.METRES + m_intLeftColumns + 12 Then Exit For 'Stop drawing behind the finish line
                For k = objRace.NUMBER_ENROLLED * 2 + 9 + m_intTopRows To objRace.NUMBER_ENROLLED * 2 + 19 + m_intTopRows
                    Cells(k, j).Interior.color = g_wksPIC.Cells(z, m_arr_varAdv(i)).Value
                    z = z + 1
                Next k 'Next row
            Next j 'Next column
            advPos = advPos + g_wksPIC.Cells(2, m_arr_varAdv(i)) 'Column for the beginning of the next ad
        Next i
    End If
    
    'Place the cursor far away
    Cells(100 + m_intTopRows, 1).Select
    
    'Activate screen updating
    Application.ScreenUpdating = True
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "DrawRaceTrack()")
    Call basAuxiliary.CodeCrash
End Sub

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
        ReDim m_arr_lngLugworms(1 To (g_wksPIC.Cells(rows.count, c).End(xlUp).row - 1)) 'ToDo... geht das in AI??? TEST in 150.00
        
        'Read the character codes for the lugworm characters from the worksheet "PIC"
        For i = 1 To UBound(m_arr_lngLugworms)
            m_arr_lngLugworms(i) = g_wksPIC.Cells(i + 1, c)
        Next i

        For i = (5 + m_intTopRows) To (objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows) 'Loop through the rows
            For j = m_intLeftColumns + 15 To objRace.METRES + m_intLeftColumns + 7 'Loop through the columns
                intLugwormFrequency = Int(((100 / objOption.LUGWORMS) - 1 + 1) * Rnd + 1) 'Lugworm or no lugworm
                If intLugwormFrequency = 1 Then
                    intLugwormShape = Int((UBound(m_arr_lngLugworms) - 1 + 1) * Rnd + 1) 'Shape of the lugworm
                    
                    With Cells(i, j) 'Draw the lugworm
                        .Font.color = objOption.LUGWORM_COL 'Colour
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
        
        For i = (5 + m_intTopRows) To (objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows) 'Loop through the rows
            For j = m_intLeftColumns + 15 To objRace.METRES + m_intLeftColumns + 7 'Loop through the columns
                intPuddleFrequency = Int(((100 / objOption.TIDE) - 1 + 1) * Rnd + 1) 'Puddle frequency
                
                If intPuddleFrequency = 1 Then
                    intPuddleLength = Int((10 - 1 + 1) * Rnd + 1) 'Puddle length (number of columns)
                    intPuddleWidth = Int((2 - 1 + 1) * Rnd + 1) 'Puddle width (number of rows)
                    
                    'Draw the puddle
                    With Range(Cells(i, j), Cells(i + intPuddleWidth - 1, j + intPuddleLength - 1))
                        .Interior.color = objOption.PUDDLE_COL
                        .Font.color = objOption.PUDDLE_COL
                        .Value = "|" 'Cell content that marks the cell as a puddle (for technical purposes). Not visible as the font colour matches the cell colour
                    End With
                    
                    #If Debugging Then
                        Range(Cells(i, j), Cells(i + intPuddleWidth - 1, j + intPuddleLength - 1)) _
                            .Font.color = vbBlack 'Make the vertical bar characters visible
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
                Cells(g_arr_varHorses(j, 3), objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = _
                    g_arr_varHorses(j, 1) & " (#" & g_arr_varHorses(j, 11) & ")" 'Number and name
            End If
        Next j
    End If
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "DrawHorseNames()")
    Call basAuxiliary.CodeCrash
End Sub

'Prepare the race sheet for displaying the current speed of each horse
Private Sub MomentumFormattings()
    
    #If Debugging Then
        'Expand the column width to see the last speed values
        Range(Columns(2), Columns(9)).ColumnWidth = 4
    #End If
    
    Columns(10).ColumnWidth = 12
    With Cells(5 + m_intTopRows, 10)
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Font.color = objRace.TRACK_COLOUR 'Hide the text
        .Value = GetText(g_arr_Text, "RACEOPT052") & " (" & GetText(g_arr_Text, "RACEOPT055") & ")"
    End With
    Range(Cells(6 + m_intTopRows, 10), Cells(objRace.NUMBER_ENROLLED * 2 + 6 + m_intTopRows, 10)).Font.color = objRace.TRACK_COLOUR
    'Speed bar
    With Range(Cells(6 + m_intTopRows, 10), Cells(objRace.NUMBER_ENROLLED * 2 + 6 + m_intTopRows, 10))
        .HorizontalAlignment = xlRight 'Text alignment
        .FormatConditions.AddDatabar
    End With
End Sub

'Race information pop-up
Private Sub RaceInfoPopup()
    With frmRaceInfo
        'Place the pop-up in the upper left corner
            .StartUpPosition = 0
            .top = ActiveWindow.top + 20
            .left = ActiveWindow.left + 20
        .BackColor = objOption.RACE_INFO_COL_B
        .caption = GetText(g_arr_Text, "USERFORM006")

        'Label with the race name and year
        With .lbl_RI1
            .BackColor = objOption.RACE_INFO_COL_B
            .ForeColor = objOption.RACE_INFO_COL_F
            .caption = objRace.RACE_NAME & " " & objRace.RACE_YEAR
            .Font.size = 12
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
'    MsgBox GetText(g_arr_Text, "RACE001") & " " & GetText(g_arr_Text, "RACE006") & " " & objRace.TRACK_LOCATION & " (" & objRace.COUNTRY & "). " & vbNewLine & vbNewLine _
'            & GetText(g_arr_Text, "RACE003") & ": " & objRace.RACE_NAME & " " & GetText(g_arr_Text, "RACE007") & " " & objRace.METRES & " " & GetText(g_arr_Text, "RACE009") & "." & vbNewLine _
'            & GetText(g_arr_Text, "RACE004") & " " & objRace.NUMBER_STARTING & " " & g_arr_Grammar(4) & ".", , g_c_tool
    
    'Compile a message text that can be used in a pop-up as well as for voice output
    Dim messagetext As String
    messagetext = GetText(g_arr_Text, "RACE001") & " " & GetText(g_arr_Text, "RACE006") & " " & objRace.TRACK_LOCATION & " (" & objRace.COUNTRY & "). " & vbNewLine & vbNewLine _
            & GetText(g_arr_Text, "RACE003") & ": " & objRace.RACE_NAME & " " & GetText(g_arr_Text, "RACE007") & " " & objRace.METRES & " " & GetText(g_arr_Text, "RACE009") & "." & vbNewLine _
            & GetText(g_arr_Text, "RACE004") & " " & objRace.NUMBER_STARTING & " " & g_arr_Grammar(4) & "."
    If objOption.SPEECH Then Call SpeechOut(messagetext) 'Voice output if selected
    Call ShowMessagePopup(g_c_tool, messagetext, enumButton.OK, vbModal) 'Show a pop-up
    
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
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "StartingGrid()")
    Call basAuxiliary.CodeCrash
End Sub

'Presentation of the horses with numbers and names
Private Sub RacePresentation()
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Call basAuxiliary.ActivateRaceSheet
    
    If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:02")) 'Delay
    
    Application.DisplayCommentIndicator = xlCommentAndIndicator 'Display comments and indicators at all times
    
    'Display a comment for each horse
    For i = 1 To objRace.NUMBER_ENROLLED
        If g_arr_varHorses(i, 0) = "START" Then
            With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4))
                'Extend the comment field if the horse is favourite or in focus
                If i = m_byteFavourite(1) And objRace.RACE_ID <> "SPACE" Then
                    If objOption.FOCUSED_RUN And g_arr_varHorses(i, 11) = objOption.FOCUSED_NR Then 'Favourite and in focus
                        .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) _
                            & " (" & GetText(g_arr_Text, "RACE011") & ") >> " & GetText(g_arr_Text, "RACE012") 'Horse number, name and "(favourite) >> in focus"
                    Else 'Favourite but not in focus
                        .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) _
                            & " (" & GetText(g_arr_Text, "RACE011") & ")" 'Horse number, name and "(favourite)"
                    End If
                ElseIf objOption.FOCUSED_RUN And g_arr_varHorses(i, 11) = objOption.FOCUSED_NR Then 'In focus but no favourite
                    .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) _
                        & " >> " & GetText(g_arr_Text, "RACE012") 'Horse number, name and ">> in focus"
                Else 'No favourite and not in focus
                    .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) 'Horse number and name
                End If
                .Comment.Shape.TextFrame.Characters.Font.size = m_intFontSize 'Font size according to the zoom level
                .Comment.Shape.TextFrame.AutoSize = True 'Resize the comment field for perfect fit
            End With
            
            'In case of a Focused Run: Highight and draw a yellow dashed frame around the horse on focus
            If objOption.FOCUSED_RUN Then
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
    If objRace.RACE_ID <> "SPACE" Then
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
    If objOption.FOCUSED_RUN Then
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
        
    'Show the horse colours on the left edge if selected in the race options
    If objOption.COLOURS_LEFT Then
        For i = 1 To objRace.NUMBER_ENROLLED
            If g_arr_varHorses(i, 0) = "START" Then
                Call PaintHorse(g_arr_varHorses(i, 3), 2, g_arr_varHorses(i, 2))
            End If
        Next i
    End If

    'Show the horse names and numbers at the start if selected in the race options
    Range(Columns(11), Columns(12)).Font.ColorIndex = 1 'Change the font colour to black
    
    'Mark the favourite on the left if selected in the race options
    If objOption.HIGHLIGHT_FAV And objRace.RACE_ID <> "SPACE" Then
        Range(Cells(g_arr_varHorses(m_byteFavourite(1), 3), 11), Cells(g_arr_varHorses(m_byteFavourite(1), 3), 12)) _
            .Interior.color = 255 'Red background
        'Show the horse with its number and name on the left (in case it does not already exist)
        Call PaintHorse(g_arr_varHorses(m_byteFavourite(1), 3), 2, g_arr_varHorses(m_byteFavourite(1), 2))
        Cells(g_arr_varHorses(m_byteFavourite(1), 3), 11).Value = "#" & g_arr_varHorses(m_byteFavourite(1), 11) 'Horse number
        Cells(g_arr_varHorses(m_byteFavourite(1), 3), 12).Value = g_arr_varHorses(m_byteFavourite(1), 1) _
                & " (" & GetText(g_arr_Text, "RACE011") & ")" 'Horse name
    End If
        
    'In case of a Focused Run: adapt the frame around the focused horse
    If objOption.FOCUSED_RUN Then
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
                    'Show the horse with its number and name on the left (in case it does not already exist)
                    Call PaintHorse(g_arr_varHorses(i, 3), 2, g_arr_varHorses(i, 2))
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
    
    'In case of displaying speed bars for each horse: show the caption
    If objOption.MOMENTUM Then Cells(5 + m_intTopRows, 10).Font.color = vbBlack
    
    'Delay before the start
     If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:04")) 'Delay

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RacePresentation()")
    Call basAuxiliary.CodeCrash
End Sub

'Start of the race
Private Sub RunRace()

    'Variable used for statistical purposes
    Dim lngLoop As Long
    
    'Variables used for race information
    Dim strMetres As String
    Dim strLeader As String
    Dim intPipesProgressBar As Integer
    
    'Variables used for water splashes
    Dim intSquirtPattern As Integer
    Dim intSquirtLength As Integer
    Dim dblSquirtColour As Double
    
    'Variables for overriding settings
    Dim noRefuse As Boolean
    Dim noTactics As Boolean
    Dim noSlipstream As Boolean
    Dim noFavourite As Boolean
    Dim yesAliens As Boolean
    
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Call basAuxiliary.ActivateRaceSheet
    
    'Override settings dependent on the selected race
    Select Case objRace.RACE_ID
        Case "SPACE"
            noRefuse = True
            noTactics = True
            noSlipstream = True
            noFavourite = True
            yesAliens = True
    End Select
    
    'Set the number of running horses equal to the number of starters
    m_intHorsesRunning = objRace.NUMBER_STARTING
    
    'One out of 100 refuses to run (if activated in the race options)
    If objOption.REFUSE_RUN And Not noRefuse Then
        Dim intRefuse As Integer
        For i = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(i, 0) = "START" Then
                Randomize
                intRefuse = Int((99 - 0 + 1) * Rnd + 0) 'Random number between 0 and 99
                If intRefuse = 0 Then
                    g_arr_varHorses(i, 0) = "REFUSED"
                    m_intHorsesRunning = m_intHorsesRunning - 1
                End If
            End If
        Next i
    End If
    
    'Prepare the race information on the worksheet (if selected in the race options)
    If objOption.RACE_INFO And objOption.RACE_INFO_WKS Then
        'Calculation of the total number of vertical bars for the progress bar
        If basAuxiliary.GetScreenWidth >= 1920 Then 'Screen resolution Full HD
            intPipesProgressBar = 30 + (9 * objOption.ZOOM_LEVEL)
        Else 'Low screen resolution
            intPipesProgressBar = 4 + ((3 + objOption.ZOOM_LEVEL) * 14)
        End If
        'Formattings
        Call basAuxiliary.RaceInfoWorksheet(objOption.RACE_INFO_COL_B, objOption.RACE_INFO_COL_F, m_intTopRows, True)
    End If
        
    'Starting procedure with gates
    If objRace.STARTING_GATE = "Y" Then
        'Hide gate numbers
        Range(Cells(7 + m_intTopRows, m_intLeftColumns + 4), Cells(5 + 2 * objRace.NUMBER_ENROLLED + m_intTopRows, m_intLeftColumns + 4)).Value = ""
        If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:04")) 'Delay
        'Open the gates
        Range(Cells(6 + m_intTopRows, m_intLeftColumns + 13), Cells(objRace.NUMBER_ENROLLED * 2 + 6 + m_intTopRows, m_intLeftColumns + 13)).Interior.color = objRace.TRACK_COLOUR
    End If
        
    If objOption.SPEECH Then Call SpeechOut(GetText(g_arr_Text, "RACE034")) 'Voice output if selected
            
    'Race information data
    strMetres = GetText(g_arr_Text, "RACE008") '"m"
    strLeader = GetText(g_arr_Text, "RACEINFO001") '"The current leader is"
    m_intPosLeader = g_arr_varHorses(1, 4) - (m_intLeftColumns + 12) 'Position zero
    m_strNameLeader = ""
    
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
        
    'Game loop for the race
    Do Until m_intPlace > m_intHorsesRunning 'solange noch nicht alle im Ziel sind
    
        lngLoop = lngLoop + 1
    
        Call basAuxiliary.ActivateRaceSheet

        'Reset the counter for horses that crossed the finish line in this loop
        m_intHorsesFinishing = 0
            
        'Re-calculation of each horse´s position
        For i = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(i, 0) = "START" Then 'Only for horses that are still running
                
                'Speed factor for this loop
                g_arr_varHorses(i, 7) = SpeedLoop()

                'For development purposes: Equalise the speed factors for all horses
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

                'Calculate the exact step width in this loop (not rounded)
                If objOption.TACTICS = False Or noTactics Then 'In case of no tactics selected in the race options
                    g_arr_varHorses(i, 8) = _
                        Round(((g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + g_arr_varHorses(i, 7)) / 3), 0)
                    #If Debugging Then
                        Debug.Print "NO TACTICS    " & g_arr_varHorses(i, 8)
                    #End If
                Else 'In case of a tactical race
                    'Calculate in which section the horse runs
                    Select Case True
                        Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) < objRace.METRES * 1 / 6 'Section 1/6
                            g_arr_varHorses(i, 8) = _
                                (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                    g_arr_varHorses(i, 7) + g_arr_varHorses(i, 26)) / 4
                            #If Debugging Then
                                Debug.Print "TACTICS       " & g_arr_varHorses(i, 26)
                            #End If
                        Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) < objRace.METRES * 2 / 6 'Section 2/6
                            g_arr_varHorses(i, 8) = _
                                (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                    g_arr_varHorses(i, 7) + g_arr_varHorses(i, 27)) / 4
                            #If Debugging Then
                                Debug.Print "TACTICS       " & g_arr_varHorses(i, 27)
                            #End If
                        Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) < objRace.METRES * 3 / 6 'Section 3/6
                            g_arr_varHorses(i, 8) = _
                                (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                    g_arr_varHorses(i, 7) + g_arr_varHorses(i, 28)) / 4
                            #If Debugging Then
                                Debug.Print "TACTICS       " & g_arr_varHorses(i, 28)
                            #End If
                        Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) < objRace.METRES * 4 / 6 'Section 4/6
                            g_arr_varHorses(i, 8) = _
                                (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                    g_arr_varHorses(i, 7) + g_arr_varHorses(i, 29)) / 4
                            #If Debugging Then
                                Debug.Print "TACTICS       " & g_arr_varHorses(i, 29)
                            #End If
                        Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) < objRace.METRES * 5 / 6 'Section 5/6
                            g_arr_varHorses(i, 8) = _
                                (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                    g_arr_varHorses(i, 7) + g_arr_varHorses(i, 30)) / 4
                            #If Debugging Then
                                Debug.Print "TACTICS       " & g_arr_varHorses(i, 30)
                            #End If
                        Case Else 'Section 6/6
                            g_arr_varHorses(i, 8) = _
                                (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                    g_arr_varHorses(i, 7) + g_arr_varHorses(i, 31)) / 4
                            #If Debugging Then
                                Debug.Print "TACTICS      " & g_arr_varHorses(i, 31)
                            #End If
                        End Select
                        
                        #If Debugging Then
                            Debug.Print " >AVG B/F/L/T " & g_arr_varHorses(i, 8)
                        #End If
                End If
                    
                'Remove water splashes
                If objRace.SQUIRT = True Then
                        Range(Cells(g_arr_varHorses(i, 3) - 1, g_arr_varHorses(i, 4) - 6), _
                            Cells(g_arr_varHorses(i, 3) + 1, g_arr_varHorses(i, 4) - 14)).Interior.Pattern = xlSolid
                End If
                        
                'Remove slipstream illustration
                If objOption.SLIPSTREAM And objOption.SLIPSTREAM_SHOW And g_arr_varHorses(i, 22) > 0 _
                    And Not noSlipstream Then
                        If g_arr_varHorses(i, 4) <= objRace.METRES + m_intLeftColumns + 9 Then
                            Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9), _
                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 12)).Interior.Pattern = xlSolid
                        End If
                End If
                    
                'Reset slipstream value
                g_arr_varHorses(i, 22) = 0
                
                'Calculate slipstream value (if activated in the race options)
                If objOption.SLIPSTREAM And Not noSlipstream Then
                    For k = 1 To UBound(g_arr_varHorses) 'Loop through the horses
                        'Find an adjacent horse
                        If g_arr_varHorses(i, 15) - 1 = g_arr_varHorses(k, 15) _
                            Or g_arr_varHorses(i, 15) + 1 = g_arr_varHorses(k, 15) Then 'One row above or below
                                'Check the distance to the horse and decide whether slipstream is given
                                If g_arr_varHorses(i, 4) > g_arr_varHorses(k, 4) - 8 _
                                    And g_arr_varHorses(i, 4) < g_arr_varHorses(k, 4) - 4 Then
                                        'Determine the multiplication factor
                                        If objOption.SLIPSTREAM_DBL Then
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
    
                'Take the slipstream effect into account
                g_arr_varHorses(i, 8) = g_arr_varHorses(i, 8) + (g_arr_varHorses(i, 22) * 100)
                
                #If Debugging Then
                    If objOption.SLIPSTREAM And g_arr_varHorses(i, 22) > 0 _
                        And Not noSlipstream Then
                            Debug.Print "+SLIPSTREAM   " & g_arr_varHorses(i, 22) * 100
                            Debug.Print "INCL SLIPSTR  " & g_arr_varHorses(i, 8)
                    End If
                #End If
    
                'Multiply position with the race speed factor
                g_arr_varHorses(i, 8) = g_arr_varHorses(i, 8) * objOption.SPEED_FACTOR
                    
                #If Debugging Then
                    Debug.Print "*SPEEDFACTOR  " & objOption.SPEED_FACTOR
                    Debug.Print " >INCL SPEEDF " & g_arr_varHorses(i, 8)
                    Debug.Print "POSITION OLD >> " & g_arr_varHorses(i, 9)
                    Debug.Print "COLUMN OLD >> " & g_arr_varHorses(i, 4)
                    Debug.Print "STEP WIDTH   >> " & g_arr_varHorses(i, 8)
                    Debug.Print "POSITION NEW >> " & g_arr_varHorses(i, 9) + g_arr_varHorses(i, 8)
                    Debug.Print "COLUMN NEW >> " & m_intLeftColumns + 12 + Round((g_arr_varHorses(i, 9) + g_arr_varHorses(i, 8)) / 1000, 0)
                    Debug.Print "METRES RUN >> " & Format(Round((g_arr_varHorses(i, 9) + g_arr_varHorses(i, 8)) / 1000, 2), "0.00")
                #End If

                'Calculate and display the momentum
                If objOption.MOMENTUM Then
                    If g_arr_varHorses(i, 0) = "START" Then
                        Cells(g_arr_varHorses(i, 3), lngLoop Mod 8 + 2).Value = g_arr_varHorses(i, 8) 'Latest speed value
                        Cells(g_arr_varHorses(i, 3), 10).Value = Round(WorksheetFunction.Average(Range(Cells(g_arr_varHorses(i, 3), 2), Cells(g_arr_varHorses(i, 3), 9))), 0)
                    End If
                End If

            End If
        Next i 'End of the re-calculation of the positions
        
        #If Debugging Then
            Debug.Print
        #End If
            
        'Horses are running
        For i = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(i, 0) = "START" Then 'Only for horses that are still running
            
                Call basAuxiliary.ActivateRaceSheet
                                        
                'Delete the horse on the worksheet
                'by assigning the track colour to the horse´s position
                Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
                    Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)) _
                    .Interior.color = objRace.TRACK_COLOUR

                'In case of a aliens around: Check the alien behaviour
                If yesAliens Then
                    If objOption.SPACE_ALIENS = enumAliens.unfriendly Then
                        Dim lngKidnapping As Long
                        Randomize
                        lngKidnapping = Int(((20000 / objOption.SPACE_KIDNAPPINGRATE) - 0 + 1) * Rnd + 0) 'Random number between 0 and 20000
                        If lngKidnapping = 0 Then
                            g_arr_varHorses(i, 0) = "KIDNAPPED"
                            If objOption.MOMENTUM Then Cells(g_arr_varHorses(i, 3), 10).ClearContents 'Delete speed value
                            Cells(g_arr_varHorses(i, 3), 12).Value = Cells(g_arr_varHorses(i, 3), 12).Value & " >>> " & GetText(g_arr_Text, "RACESPEC014") '" >>> kidnapped"
                            m_intHorsesRunning = m_intHorsesRunning - 1
                            Exit For
                        End If
                    End If
                End If
                            
                'Calculate the new position of the horse
                g_arr_varHorses(i, 9) = g_arr_varHorses(i, 9) + g_arr_varHorses(i, 8) 'Exact (internal) position (accuracy 0.25 millimetres)
                'Make sure that no horse moves behind the starting line
                If g_arr_varHorses(i, 9) < 0 Then g_arr_varHorses(i, 9) = 0
                'Calculate the position on the worksheet (accuracy 1 metre)
                g_arr_varHorses(i, 4) = m_intLeftColumns + 12 + Round(g_arr_varHorses(i, 9) / 1000, 0)
                        
                'Draw slipstream illustration (if selected in the race options)
                If objOption.SLIPSTREAM And objOption.SLIPSTREAM_SHOW And g_arr_varHorses(i, 22) > 0 _
                    And g_arr_varHorses(i, 4) <= objRace.METRES + m_intLeftColumns + 7 _
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
                            Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9).Interior.color = 14395790
                        Else 'Double slipstream effect: dark blue
                            Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9).Interior.color = 9851952
                        End If
                    #End If
                End If
                        
                'Draw water splashes
                If objRace.SQUIRT = True And g_arr_varHorses(i, 4) > 13 + m_intLeftColumns _
                    And g_arr_varHorses(i, 4) <= objRace.METRES + m_intLeftColumns Then
                    
                    If Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)).Interior.color = objOption.PUDDLE_COL Then
                    
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
                    For j = 0 To 7 'Loop through each segment of the horse
                        If IsArray(g_arr_varHorses(i, 2)) Then
                            Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7 + j) _
                                .Font.color = g_arr_varHorses(i, 2)(j) 'Font colour
                        Else
                            Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7 + j) _
                                .Font.color = g_arr_varHorses(i, 2) 'Font colour
                        End If
                    Next j
                    
                    'Show the lugworms again behind the horse
                    'by assigning the lugworm colour
                    Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8), _
                        Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - 2 * objOption.SPEED_FACTOR)) _
                            .Font.color = objOption.LUGWORM_COL 'Font colour
                            
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
                                    .Font.color = objOption.PUDDLE_COL 'Font colour
                                    .Interior.color = objOption.PUDDLE_COL 'Cell colour
                                End With
                                #If Debugging Then 'Show the water indicator sign in red
                                    Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - k).Font.color = vbRed
                                #End If
                            Case Else
                                'Overpaint with the original track colour
                                With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - k)
                                    .Interior.color = objRace.TRACK_COLOUR 'Cell colour
                                End With
                        End Select
                    Next k
        
                End If
                    
                'SPECIAL: Wild boar track devastation illustration
                If g_arr_varHorses(i, 24) = "WILD" _
                     And g_arr_varHorses(i, 4) < objRace.METRES + m_intLeftColumns + 12 Then
                     'Draw a race track devastation sign (#) under the left segment of the wild boar
                     If IsArray(g_arr_varHorses(i, 2)) Then 'Multicoloured wild boar
                        With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)
                            .Font.color = g_arr_varHorses(i, 2)(0) 'Colour of the left segment
                            .Value = "#"
                        End With
                    Else 'Monochrome wild boar
                        With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)
                            .Font.color = g_arr_varHorses(i, 2) 'Wild boar colour
                            .Value = "#"
                        End With
                    End If
                End If
    
                'Draw hoof prints behind the horse (if selected in the race options)
                If objOption.HOOFPRINTS And IsEmpty(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8)) _
                    And g_arr_varHorses(i, 4) > 13 + m_intLeftColumns _
                    And Not g_arr_varHorses(i, 24) = "WILD" Then _
                    Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8).Value = "'-" 'Hoof print character
            
            End If
            
            'Horizontal scrolling dependent on the camera mode
            If objOption.FOCUSED_RUN Then 'Scrolling in a Focused Run
                'Find the horse in focus
                If g_arr_varHorses(i, 11) = objOption.FOCUSED_NR And g_arr_varHorses(i, 0) = "START" Then
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
            Else 'Scrolling in standard mode
                'Check whether the leading horse is near the right edge of the screen
                If g_arr_varHorses(i, 4) > ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column _
                            - ((ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column _
                            - ActiveWindow.VisibleRange.Column) * 1 / 10) _
                    And ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column <= objRace.METRES + m_intLeftColumns + 2 Then
                        'Standard scrolling in the style of paging
                        ActiveWindow.ScrollColumn = ActiveWindow.VisibleRange.Column _
                                            + ((ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).Column _
                                            - ActiveWindow.VisibleRange.Column) * 8 / 10)
                End If
            End If
            
            'Show race information (if selected in the race options)
            If objOption.RACE_INFO Then
                'Check the data of the current leader
                If (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) > m_intPosLeader Then
                    m_intPosLeader = g_arr_varHorses(i, 4) - m_intLeftColumns - 12 'Position of the leader
                    m_strNameLeader = g_arr_varHorses(i, 1) 'Name of the leader
                End If
                
                'Refresh the race information data (on the worksheet)
                If objOption.RACE_INFO_WKS Then
                    'Race distance progress bar
                    If objOption.RACE_INFO_PROGRESS Then
                        With Cells(3 + m_intTopRows, 11)
                            'Relation of metres run and the total race distance
                            .Value = m_intPosLeader & strMetres & " / " & objRace.METRES & strMetres
                        End With
                        If (intPipesProgressBar / objRace.METRES * m_intPosLeader) < 1 Then
                            'Show the first bar as soon as the race is started
                            Cells(3 + m_intTopRows, 12).Value = "|"
                        Else
                            With Cells(3 + m_intTopRows, 12)
                                'Calculate the number of bars depending on the leader´s position
                                .Value = String((intPipesProgressBar / objRace.METRES * m_intPosLeader), "|")
                            End With
                        End If
                    End If
                    'Name of the leader
                    If objOption.RACE_INFO_LEADER Then
                        'Only in the range between 20 meters after the start and 20 meters before the finish
                        If m_intPosLeader > 20 And m_intPosLeader < (objRace.METRES - 20) Then
                            With Cells(1 + m_intTopRows, 2)
                                .Value = strLeader
                            End With
                            With Cells(2 + m_intTopRows, 10)
                                .Value = m_strNameLeader
                            End With
                        Else
                            'For performance reasons the labels must be provided with a space character
                            Cells(1 + m_intTopRows, 2).Value = " "
                            Cells(2 + m_intTopRows, 10).Value = " "
                            'Comment in the following lines for experiencing
                            'the speed degradation while no leader is displayed
'                            Cells(1 + m_intTopRows, 2).Value = ""
'                            Cells(2 + m_intTopRows, 10).Value = ""
                        End If
                    End If
                End If
                
                'Refresh the race information data (in a pop-up)
                If objOption.RACE_INFO_POP Then
                    'Race distance progress bar
                    If objOption.RACE_INFO_PROGRESS Then
                        With frmRaceInfo.Controls("lbl_RI3b_dyn")
                            .width = CInt(200 / objRace.METRES * m_intPosLeader)
                            .caption = m_intPosLeader
                        End With
                    End If
                    'Name of the leader
                    If objOption.RACE_INFO_LEADER Then
                        'Only in the range between 20 meters after the start and 20 meters before the finish
                        If m_intPosLeader > 20 And m_intPosLeader < (objRace.METRES - 20) Then
                            frmRaceInfo.Controls("lbl_RI4a_dyn").caption = strLeader
                            frmRaceInfo.Controls("lbl_RI4b_dyn").caption = m_strNameLeader
                        Else 'Dont´show the leader´s name in the first and the last 20 meters
                            frmRaceInfo.Controls("lbl_RI4a_dyn").caption = " "
                            frmRaceInfo.Controls("lbl_RI4b_dyn").caption = " "
                        End If
                    End If
                End If
            End If
        Next i
            
        'Check whether one or more horses have reached the finish line
        For i = 1 To UBound(g_arr_varHorses) 'Loop through the horses
            If g_arr_varHorses(i, 0) = "START" Then 'Only for horses that are still running
            
                #If Debugging Then 'For debugging purposes: Name and position (accuracy 10cm)
                    Debug.Print "#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) & " - Position: " _
                        & Format(Round(g_arr_varHorses(i, 9) / 1000, 2), "0.00") & " metres"
                #End If
            
                If g_arr_varHorses(i, 9) >= objRace.METRES * 1000 Then 'Horse has crossed the finish line
                    #If Debugging Then
                        Debug.Print "  >> FINISHED - Exact position: " & g_arr_varHorses(i, 9)
                    #End If
                    g_arr_varHorses(i, 0) = "CALCULATION" 'Set the horse´s status
                    m_intHorsesFinishing = m_intHorsesFinishing + 1 'Count the number of horses that pass the finish line in this loop
                    If objOption.MOMENTUM Then Cells(g_arr_varHorses(i, 3), 10).ClearContents
                End If
            End If
        Next i
        
        'Evaluate the number of finishers in this loop
        If m_intHorsesFinishing > 0 Then
            If m_blnWin = False Then 'If there is no winner yet
                If m_intHorsesFinishing > 1 Then
                    m_blnPhotofinish = True 'In case of more than one possible winners: Flag for a tight finish
                End If
                Call CreateFinishPhoto 'Create a photo of the finish
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
        
    'Remove race information
    If objOption.RACE_INFO Then
        If objOption.RACE_INFO_POP Then Unload frmRaceInfo 'Close the pop-up
        If objOption.RACE_INFO_WKS Then Call basAuxiliary.RaceInfoWorksheet(xlNone, 0, m_intTopRows, False) 'Delete on the worksheet
    End If
    
    'Remove race speed caption
    If objOption.MOMENTUM Then Cells(5 + m_intTopRows, 10).ClearContents
    
    'In case of a photo finish
    If m_blnPhotofinish = True Then
        Call basAuxiliary.ActivateRaceSheet
        If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:04")) 'Delay
        'Clear text "PHOTO FINISH!"
        Cells(2 + m_intTopRows, objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Clear
        'Unfreeze the window pane if it is frozen
        If objOption.NAMES_LEFT Or objOption.COLOURS_LEFT Or objOption.HIGHLIGHT_FAV _
            Or (objOption.FOCUSED_RUN And objOption.HIGHLIGHT_FOC) _
            Or (objOption.RACE_INFO And objOption.RACE_INFO_WKS) _
                Then Call basAuxiliary.Freeze(0, 0, False)
        'Scroll to the area where the photo will be displayed
        ActiveWindow.ScrollColumn = objRace.METRES + m_intLeftColumns + 15 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)
        'Black photo background
        Range(Cells(5 + m_intTopRows, objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
            Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METRES + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))).Interior.ColorIndex = 1
        'Display text "Photo creation"
        Cells(4 + m_intTopRows, objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = GetText(g_arr_Text, "RACE026")
        If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:04")) 'Delay
        'Show the photo of the tight finish
        Call DrawFinishPhoto
        'Display text "Photo evaluation"
        Cells(4 + m_intTopRows, objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = GetText(g_arr_Text, "RACE027")
    End If
    
    If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:04")) 'Delay
    'Clear text above the photo
    Cells(4 + m_intTopRows, objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).ClearContents

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RunRace()")
    Call basAuxiliary.CodeCrash
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
        If g_arr_varHorses(i, 0) = "CALCULATION" Then
            g_arr_varHorses(i, 0) = "FINISHED" 'Set the horse´s new status
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
        If g_arr_varHorses(i, 0) = "REFUSED" Or g_arr_varHorses(i, 0) = "KIDNAPPED" Then 'Find the horses that refused to run
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

'Formattings for the photo of the finish and the ranking list
Private Sub FormatPhotoAndRanking()

    'Texts in case of a photo finish
    Cells(2 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)) _
        .Font.name = "Arial Black" '"PHOTO FINISH!"
    With Cells(4 + m_intTopRows, _
            objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)) '"Photo..."
        .Font.size = 14
        .Font.Bold = True
    End With

    'Column width
    Range(Columns(objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
        Columns(objRace.METRES + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))) _
        .ColumnWidth = m_dblRankingsWidth / 10 'Columns for the photo of the finish
    Columns(objRace.METRES + m_intLeftColumns + 176 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR)) _
        .ColumnWidth = m_dblRankingsWidth   'Column behind the photo of the finish

End Sub

'Create a photo of the finish when the first horse
'crosses the finish line
Private Sub CreateFinishPhoto()
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

            'Announcement "PHOTO FINISH!"
            If objOption.SPEECH Then Call SpeechOut(GetText(g_arr_Text, "RACE025")) 'Voice output if selected
            Cells(2 + m_intTopRows, objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = GetText(g_arr_Text, "RACE025")
            
            'Alternate the track colour rapidly between black and white behind the finish line
            For k = 1 To 8 'Run the loop 8 times
                With Range(Cells(5 + m_intTopRows, objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
                    Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)))
                        .Interior.ColorIndex = 1 'Black
                        .Interior.ColorIndex = 0 'White
                End With
            Next k
            'Reset the track to its original colour
            Range(Cells(5 + m_intTopRows, objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
                Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METRES + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))).Interior.color = objRace.TRACK_COLOUR
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
        colourTrack = 1  'Black track
        colourLines = 2 'White lines
        colourNames = vbWhite
        colourScale = vbWhite
    Else
        colourTrack = objRace.TRACK_COLOUR 'Original track colour
        colourLines = 1 'Black lines
        colourNames = vbBlack
        colourScale = objRace.TRACK_COLOUR
    End If
    
    Application.ScreenUpdating = False 'Deactivate screen updating
    
    'Clear the range on which the photo is to be shown
    Range(Cells(5 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 8 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))) _
            .Clear
    
    'Draw a frame around the photo
    Range(Cells(5 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 8 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))) _
            .BorderAround ColorIndex:=0, Weight:=xlMedium
            
    'Draw a horizontal line for the section with the scale markers
    With Range(Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))) _
            .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .ColorIndex = colourLines
                .Weight = xlThin
    End With
    
    'Write the race data as a caption
    With Cells(5 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))
            .Font.name = "Arial"
            .Font.Bold = True
            .Font.color = colourNames
            .Value = objRace.RACE_NAME & " " & objRace.RACE_YEAR _
                & " - " & objRace.TRACK_NAME & ", " & objRace.TRACK_LOCATION _
                & " - " & objRace.METRES & " " & GetText(g_arr_Text, "RACE009")
    End With

    'Draw the race track and the finish line
    Range(Cells(5 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))) _
            .Interior.color = colourTrack 'Background
    Range(Cells(objRace.NUMBER_ENROLLED * 2 + 8 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 8 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))) _
            .Interior.color = colourScale 'Stripe for the scale markers
    Range(Cells(5 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 155 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
        Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 155 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))) _
        .Interior.ColorIndex = colourLines 'Finish line
        
    'Draw track markers ("XXX metres")
    For i = 1 To (15 + 2 * objOption.SPEED_FACTOR)
        Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, _
            i * 10 + objRace.METRES + m_intLeftColumns + 15 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)) _
            .Interior.ColorIndex = colourLines 'Scale (vertical bars)
        With Cells(objRace.NUMBER_ENROLLED * 2 + 8 + m_intTopRows, _
            i * 10 + objRace.METRES + m_intLeftColumns + 13 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))
            .Value = (objRace.METRES - 14 + i) & GetText(g_arr_Text, "RACE008")
            .Font.color = vbBlack 'Scale (metres)
        End With
    Next i

    'Prepare variables
    Dim photoLeftMargin As Integer 'Column of the left edge of the photo
    Dim racetrackMinimumColumn As Long 'Minimum horse head position to appear on the photo
    Dim currentDrawColumn As Integer 'Column for drawing the current segment

    photoLeftMargin = objRace.METRES + m_intLeftColumns + 16 _
        + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR) 'Get the left column of the photo
    racetrackMinimumColumn = objRace.METRES * 10 - 139 'Get the minimum position to appear
                
    'Paint the horses
    For i = 1 To UBound(m_arr_varPhotofinish)
        If m_arr_varPhotofinish(i, 1) >= racetrackMinimumColumn Then 'Only if the horse appears in the photo (at least partially)
            currentDrawColumn = photoLeftMargin + m_arr_varPhotofinish(i, 1) - racetrackMinimumColumn 'Column of the horse´s head
          
            'Find the next horse to paint (sort by horse numbers ascending)
            For j = 1 To UBound(m_arr_varResults)
                If m_arr_varResults(j, 2) = m_arr_varPhotofinish(i, 2) Then
                    Exit For
                End If
            Next j
            
            'Draw a vertical line at the position of the horse´s head
            '(only for those which have crossed the finish line)
            If m_arr_varPhotofinish(i, 1) >= (objRace.METRES * 10) Then
                With Range(Cells(5 + m_intTopRows, currentDrawColumn), _
                    Cells(objRace.NUMBER_ENROLLED * 2 + 8 + m_intTopRows, currentDrawColumn)) _
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
                          .Interior.color = horseSegmentColour
                    
                    currentDrawColumn = currentDrawColumn - segmentLength 'Adjust the column for the next segment
                End If
            Next horseSegment 'Next segment
        End If
        
        'Write the horse names in the photo (if selected in the race options)
        If objOption.NAMES_PHOTO = True And g_arr_varHorses(i, 0) <> "CANCELLED" Then
            With Cells(g_arr_varHorses(i, 3), _
                    objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))
                .Font.name = "Arial"
                .Font.size = m_intFontSize
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
                .Font.color = colourNames
                .Value = g_arr_varHorses(i, 1) 'Horse name
            End With
        End If
    Next i

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
            Or (objOption.FOCUSED_RUN And objOption.HIGHLIGHT_FOC) Or (objOption.RACE_INFO And objOption.RACE_INFO_WKS) Then Call basAuxiliary.Freeze(0, 0, False)
    'Scrollen zu Ergebnistafel
        Call basAuxiliary.Scroll(objRace.METRES + m_intLeftColumns + 15 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR), m_intTopRows + (objRace.NUMBER_ENROLLED * 2 + 9))

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "RaceFinished()")
    Call basAuxiliary.CodeCrash
End Sub

'Draw the ranking list
Private Sub DrawRankingList(afterRace As Boolean)
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Call basAuxiliary.ActivateRaceSheet

    'Scoreboard
    With Range(Cells(objRace.NUMBER_ENROLLED * 2 + 20 + m_intTopRows, _
            objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
            Cells(objRace.NUMBER_ENROLLED * 2 + 20 + objRace.NUMBER_STARTING + 1 + m_intTopRows, _
            objRace.METRES + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)))
        .Clear 'clear all cell values and formattings
        .BorderAround ColorIndex:=0, Weight:=xlMedium 'Border
        .Interior.color = 16777215 'Background
        .Font.name = "Courier New"
        .Font.size = 12
       .NumberFormat = "@" 'Force text format
    End With
    With Cells(objRace.NUMBER_ENROLLED * 2 + 20 + m_intTopRows, _
                objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))
        .Font.size = 14 'Formattings for the headline
        .Font.Bold = True
        .IndentLevel = 1
    End With
    
    'Headline
    Cells(objRace.NUMBER_ENROLLED * 2 + 20 + m_intTopRows, _
            objRace.METRES + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)) _
            .Value = objRace.RACE_NAME & " " & objRace.RACE_YEAR & " - " & objRace.TRACK_LOCATION
            
    If Not g_skipDelay And objOption.RANKING_DELAY And afterRace Then _
        Application.Wait (Now + TimeValue("0:00:04")) 'Delay
        
    'Results
    Dim intPositionName As Integer 'Position of the horse names
    intPositionName = 0
    For i = UBound(m_arr_varResults) To 1 Step -1 'Loop through through the results bottom up
        If objOption.RANKING_COL Then 'Show colours if selected in the race options
            intPositionName = 12
            Call PaintHorse(objRace.NUMBER_ENROLLED * 2 + 20 + i + m_intTopRows, _
                objRace.METRES + m_intLeftColumns + 19 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR), _
                m_arr_varResults(i, 4)) 'Paint the horse
            Range(Cells(objRace.NUMBER_ENROLLED * 2 + 20 + i + m_intTopRows, _
                objRace.METRES + m_intLeftColumns + 19 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
                Cells(objRace.NUMBER_ENROLLED * 2 + 20 + i + m_intTopRows, _
                objRace.METRES + m_intLeftColumns + 26 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))) _
                .BorderAround ColorIndex:=0, Weight:=xlThin 'Draw a frame around the horse
        End If
        Cells(objRace.NUMBER_ENROLLED * 2 + 20 + i + m_intTopRows, _
            objRace.METRES + m_intLeftColumns + 22 + intPositionName + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)) _
            .Value = m_arr_varResults(i, 1) & "."  'Place
        Cells(objRace.NUMBER_ENROLLED * 2 + 20 + i + m_intTopRows, _
            objRace.METRES + m_intLeftColumns + 29 + intPositionName + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)) _
            .Value = m_arr_varResults(i, 3) & " (#" & m_arr_varResults(i, 2) & ")" 'Horse name and number
        If Not g_skipDelay And objOption.RANKING_DELAY And afterRace Then _
            Application.Wait (Now + TimeValue("0:00:01")) 'Delay
    Next i
                
    'Show a pop-up in case of a dead heat (more than one winner)
    If m_blnDeadHeat Then ShowDeadHeat
    'Alternative variants:
'    If m_blnDeadHeat = True Then Call ShowDeadHeat
'    If m_blnDeadHeat = True Then ShowDeadHeat
'    If m_blnDeadHeat Then Call ShowDeadHeat

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "DrawRankingList()")
    Call basAuxiliary.CodeCrash
End Sub

'Speed factor for a single step (recalculated for each horse in every loop)
Function SpeedLoop() As Double
    Randomize 'Initialize the random number generator
    'Get a random number within a defined range
    SpeedLoop = Int((objSpeed.SPEED_LOOP_HIGH - objSpeed.SPEED_LOOP_LOW + 1) * Rnd + objSpeed.SPEED_LOOP_LOW)
End Function

'Display the winner with a wreath
Private Sub DrawWinnerPhoto()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Call basAuxiliary.ActivateRaceSheet
    
    'Formattings for the name of the winner
    With Range(Cells(objRace.NUMBER_ENROLLED * 2 + 20 + m_intTopRows, _
            objRace.METRES + m_intLeftColumns + 177 + (2 * 10 * objOption.SPEED_FACTOR)), _
            Cells(objRace.NUMBER_ENROLLED * 2 + 21 + m_intTopRows, _
            objRace.METRES + m_intLeftColumns + 179 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR)))
        .Font.size = 14
        .Font.Bold = True
    End With
    
    For i = 1 To UBound(m_arr_varResults)
        If m_arr_varResults(i, 1) = 1 Then

            'Column width for the photo
            Range(Columns(objRace.METRES + m_intLeftColumns + 177 + ((i - 1) * 21) + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR)), _
            Columns(objRace.METRES + m_intLeftColumns + 197 + ((i - 1) * 21) + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))) _
            .ColumnWidth = 2

            'Draw the photo
            Call PaintPicture(g_wksPIC, g_wksRace, m_arr_varResults(i, 6), _
                objRace.METRES + m_intLeftColumns + 175 + ((i - 1) * 21) + 19 + 2 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR), _
                m_intTopRows + objRace.NUMBER_ENROLLED * 2 + 23 + 12 + 5, _
                m_intTopRows + objRace.NUMBER_ENROLLED * 2 + 23, _
                objRace.METRES + m_intLeftColumns + 177 + ((i - 1) * 21) + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))

            'Draw a frame around the photo
            Range(Cells(objRace.NUMBER_ENROLLED * 2 + 23 + m_intTopRows, _
                objRace.METRES + m_intLeftColumns + 177 + ((i - 1) * 21) + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR)), _
                Cells(objRace.NUMBER_ENROLLED * 2 + 40 + m_intTopRows, _
                objRace.METRES + m_intLeftColumns + 196 + ((i - 1) * 21) + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))) _
                .BorderAround ColorIndex:=0, Weight:=xlThin 'Draw a frame around the horse
        
        End If
    Next i
    
    Cells(objRace.NUMBER_ENROLLED * 2 + 20 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 177 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR)) _
        .Value = GetText(g_arr_Text, "RACE031") '"Winner of the race:"
    Cells(objRace.NUMBER_ENROLLED * 2 + 21 + m_intTopRows, _
        objRace.METRES + m_intLeftColumns + 179 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR)) _
        .Value = m_strWinner 'Name of the winner
    
    'Voice output if selected in the race options
    If objOption.SPEECH Then Call SpeechOut(GetText(g_arr_Text, "RACE031"))
    If objOption.SPEECH Then Call SpeechOut(m_strWinner)
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "DrawWinnerPhoto()")
    Call basAuxiliary.CodeCrash
End Sub

'Check for a dead heat (i.e. more than one winner)
Private Sub CheckDeadHeat()
    'Reset variables
    m_blnDeadHeat = False
    m_strWinner = ""
    
    'Loop through the array with the race results
    For i = 1 To UBound(m_arr_varResults)
        If m_arr_varResults(i, 1) = 1 Then 'Find horses that rank 1st
            If i > 1 Then
                'In case of more than one winner
                m_strWinner = m_strWinner & " " & GetText(g_arr_Text, "RACE017") & " " 'add " and "
                m_blnDeadHeat = True 'Flag for a dead heat
            End If
            'Compile the string with the name(s) of the winner(s)
            m_strWinner = m_strWinner & UCase(m_arr_varResults(i, 3))
        End If
    Next i
    
End Sub

'Show a pop-up in case of a dead heat
Private Sub ShowDeadHeat()
    Call ShowInfoPopup(objRace.RACE_NAME & " " & objRace.RACE_YEAR, _
        UCase(GetText(g_arr_Text, "RACE033")) & "!", _
        False, vbModal, 22)
End Sub

'Analyse bettings (currently only for type "win")
Private Sub AnalyseBettings()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    'Variables for a single bet slip
    Dim id As String
    Dim nm As String
    Dim ty As String
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
            payout = True 'Set the initial value for pay-out
            'Get the data of a bet
            id = g_colBetSlips(i).id
            nm = g_colBetSlips(i).GamblerName
            ty = g_colBetSlips(i).BType
            St = g_colBetSlips(i).Stake
            od = g_colBetSlips(i).Odds * 10
            bt() = g_colBetSlips(i).bet 'Prediction of the ranking (place 1-4, dependent on the type of the bet)
            
            'Analyse the bet slip by comparing the prediction with the actual rank
            For j = 1 To UBound(bt) 'Loop through the array with the prediction
                If bt(j) <> m_arr_varResults(j, 2) Then payout = False
            Next j
            
            'Set the value and the colour of the pay-out label
            If payout = True And m_blnDeadHeat = False Then 'No payout in case of a dead heat
                payCash = St / 10 * od 'Pay-out
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
                .caption = UCase(GetText(g_arr_Text, "BET007")) & ": " & ty '"TYPE OF BET:"
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
                    .caption = GetText(g_arr_Text, "BET041") & ": " & GetHorseName(bt(j)) & " (#" & bt(j) & ")" '"To win:"
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
                .caption = "  " & GetText(g_arr_Text, "BET038") & ": " & Format(payCash, "0.00") & " " & GetText(g_arr_Text, "BET035") 'Pay-out: xx EUR
                .BackColor = payColor
            End With
                
            'For statistical purposes: Calculate the total stakes and pay-out
            totalStake = totalStake + St
            totalPayout = totalPayout + payCash
            
            align = align + 2
            
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
            .width = 300
            .TextAlign = fmTextAlignLeft
            .caption = GetText(g_arr_Text, "BET043") & ": " & Format(totalStake, "0.00") & " " & GetText(g_arr_Text, "BET035") & " / " _
                        & GetText(g_arr_Text, "BET044") & ": " & Format(totalPayout, "0.00") & " " & GetText(g_arr_Text, "BET035")
        End With

        'Pop-up formattings
        With frmBettingAnalysis
            .caption = objRace.RACE_NAME & " " & objRace.RACE_YEAR & " | " & objRace.TRACK_NAME & ", " & objRace.TRACK_LOCATION _
                        & " (" & objRace.COUNTRY & ")" 'Race data
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
    End If

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "AnalyseBettings()")
    Call basAuxiliary.CodeCrash
End Sub

'Retrieve the horse name by the horse number
Private Function GetHorseName(num As Integer) As String
    Dim X As Integer
    For X = 1 To UBound(g_arr_varHorses())
        If num = g_arr_varHorses(X, 11) Then Exit For
    Next
    GetHorseName = g_arr_varHorses(X, 1)
End Function

'xxxxx Delete all GaloppSim worksheets
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
        .lblS3.caption = objRace.TRACK_NAME & " " & GetText(g_arr_Text, "RACE002") & " " & objRace.TRACK_LOCATION & " (" & objRace.COUNTRY & ")" 'Race track, loaction and country
        .lblS6.caption = objRace.NUMBER_STARTING & " " & g_arr_Grammar(4) 'Number of horses starting
        If objRace.REAL_RACE = "Y" Then
            .lblS10.caption = UCase(GetText(g_arr_Text, "START009")) '"REAL RACE"
        Else
            .lblS10.caption = UCase(GetText(g_arr_Text, "START010")) '"FICTITIOUS RACE"
        End If
        .lblS8.caption = GetText(g_arr_Text, "START005") 'Caption of the betting section
        Call GetNumberBetSlips 'Refresh the number of bet slips
        .lblFocus.caption = g_arr_Grammar(1) & " " & GetText(g_arr_Text, "START006") 'Label "Horse in focus"
        .cmdS1.caption = GetText(g_arr_Text, "START002") 'Button "Add bet slip"
        .cmdS2.caption = GetText(g_arr_Text, "START003") 'Button "Start the race"
        With .lblS4 'Track surface
            .caption = objRace.TRACK_SURFACE_TEXT 'Surface type
            .BorderStyle = fmBorderStyleSingle 'Draw a border around the track preview
            .BackColor = objRace.TRACK_COLOUR 'Set the colour according to the track colour
        End With
        .cmdS5.caption = GetText(g_arr_Text, "START014") 'button "Speed and form"
        .cmdS6.caption = GetText(g_arr_Text, "START015") 'button "Odds"
        'Height of the pop-up
        If objOption.BET_MODE = True And objRace.BETTING_ALLOWED = "Y" Then
            .Height = 315 'If the betting mode is enabled
        Else
            .Height = 200 'If the betting mode is disabled
        End If
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
        .lblC1 = objRace.TRACK_NAME & " - " & objRace.TRACK_LOCATION & " (" & objRace.COUNTRY & ")"
        .lblC2 = objRace.RACE_NAME & " " & objRace.RACE_YEAR
        .show (vbModal)
    End With
End Sub

'Pop-up with speed bars and odds
Public Sub ShowSpeed(Odds As Boolean)
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
        If Odds Then
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
                If Odds Then
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
                Dim xxx As Integer
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
    Dim bet As String, horsename As String
    Dim i As Integer, j As Integer
    Dim bt() As Integer
    bt() = g_colBetSlips(id).bet 'Get the array with the guesses
    For i = 1 To UBound(bt) 'Loop through the array
        For j = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(j, 11) = bt(i) Then 'Find the horse name
                horsename = g_arr_varHorses(j, 1)
                Exit For
            End If
        Next j
        bet = bet & bt(i) & " " & horsename & vbNewLine 'Horse number and name
    Next i

    'Write all data to the receipt
    With frmReceipt
        .caption = g_colBetSlips(id).GamblerName
        .lblR1 = UCase(objRace.TRACK_LOCATION) ' & " (" & objRace.COUNTRY & ")")
                                            '(comment in for adding the country)
        .lblR2 = UCase(objRace.RACE_NAME)
        .lblR3 = objRace.NUMBER_STARTING & " " & UCase(g_arr_Grammar(4))
        .lblR4 = UCase(g_colBetSlips(id).BType)
        .lblR5 = UCase(bet)
        .lblR6 = UCase(GetText(g_arr_Text, "BET036") & " " & Format(g_colBetSlips(id).Stake, "0.00") & " " & GetText(g_arr_Text, "BET035"))
        .lblR7 = g_colBetSlips(id).id
        .show (vbModal)
    End With
End Sub


'CALLBACKS (Excel ribbon events)
'-------------------------------

'xxxxx Callback for customUI.onLoad
Private Sub AI_GaloppSimAddinInitialize(ribbon As IRibbonUI)
    Set g_RibbonGaloppSim = ribbon
'    If g_wksTEXT Is Nothing Then Set g_wksTEXT = Table_TEXT 'Kann das RAUS ? TEST in 150.00
'    Call GetTextComponents 'Kann das RAUS ? TEST in 150.00
End Sub

'xxxxx Callback
Private Sub AI_GetLabel(control As IRibbonControl, ByRef returnedVal)
    
    'Display the start screen in AI edition when activating the GALOPPSIM menu tab the first time
    If AI_started = False Then
        AI_started = True
        Call TitleScreen
    End If
    
    Select Case control.id
        Case "group01GALOPPSIM" 'group "Settings"
            returnedVal = GetText(g_arr_Text, "AI001")
            Case "btn05GALOPPSIM" 'button "Race options"
                returnedVal = GetText(g_arr_Text, "BTN020")
            Case "btn10GALOPPSIM" 'button "Race options"
                returnedVal = GetText(g_arr_Text, "BTN001")
            Case "menu02GALOPPSIM" 'menu "Language"
                returnedVal = GetText(g_arr_Text, "BTN002")
                Case "cb01aGALOPPSIM" 'checkbox "excel mode"
                    returnedVal = GetText(g_arr_Text, "EXCELOPT010")
                Case "cb01bGALOPPSIM" 'checkbox "TV mode with Excel menu strip
                    returnedVal = GetText(g_arr_Text, "EXCELOPT011")
                Case "cb01cGALOPPSIM" 'checkbox "TV mode (full-screen)"
                    returnedVal = GetText(g_arr_Text, "EXCELOPT012")
                
        Case "group02GALOPPSIM" 'group "Race"
            returnedVal = GetText(g_arr_Text, "AI002")
            Case "btn30GALOPPSIM" 'button "Start race"
                returnedVal = GetText(g_arr_Text, getCaptionStartBtn(objOption.BET_MODE))
            Case "combo01InstalledRaces" 'combobox "Race selection"
                returnedVal = GetText(g_arr_Text, "RACE003")
            Case "btn31GALOPPSIM" 'button "Photo of the finish"
                returnedVal = GetText(g_arr_Text, "BTN004")
            Case "btn32GALOPPSIM" 'button "Ranking list"
                returnedVal = GetText(g_arr_Text, "BTN005")
            Case "btn33GALOPPSIM" 'button "Photo of the winner"
                returnedVal = GetText(g_arr_Text, "BTN006")
            Case "btn34GALOPPSIM" 'button "Betting analysis"
                returnedVal = GetText(g_arr_Text, "BTN007")
            
        Case "menu01GALOPPSIM" 'menu "Language"
            returnedVal = GetText(g_arr_Text, "LANGUAGE001")
            Case "btn01aGALOPPSIM" 'button "deutsch"
                returnedVal = GetText(g_arr_Text, "LANGUAGE002")
            Case "btn01bGALOPPSIM" 'button "english"
                returnedVal = GetText(g_arr_Text, "LANGUAGE003")
'                Case "btn01cGALOPPSIM" 'button "schwiizerdütsch"
'                    returnedVal = GetText(g_arr_Text, "LANGUAGE004")
            Case "btn01dGALOPPSIM" 'button "russian"
                returnedVal = GetText(g_arr_Text, "LANGUAGE005")
            Case "btn01eGALOPPSIM" 'button "bulgarian"
                returnedVal = GetText(g_arr_Text, "LANGUAGE006")
        Case "btn40GALOPPSIM" 'button "Info"
            returnedVal = GetText(g_arr_Text, "BTN009")
        Case "btn50GALOPPSIM" 'button "Warning"
            returnedVal = GetText(g_arr_Text, "BTN010")
        Case "btn60GALOPPSIM" 'button "Movie"
            returnedVal = GetText(g_arr_Text, "BTN011")
        Case "btn70GALOPPSIM" 'button "Close"
            returnedVal = GetText(g_arr_Text, "BTN012")
    End Select
End Sub

'xxxxx Callbacks for Tooltips
Private Sub AI_GetScreentip(control As IRibbonControl, ByRef screentip)
    Select Case control.id
        Case "btn05GALOPPSIM" 'Button "Start screen"
            screentip = GetText(g_arr_Text, "TIP012")
        Case "btn10GALOPPSIM" 'Button "Race options"
            screentip = GetText(g_arr_Text, "USERFORM001")
        Case "cb01aGALOPPSIM" 'Combobox "Excel mode"
            screentip = GetText(g_arr_Text, "EXCELOPT010")
        Case "cb01bGALOPPSIM" 'Combobox "TV mode with Excel menu strip
            screentip = GetText(g_arr_Text, "EXCELOPT011")
        Case "cb01cGALOPPSIM" 'Combobox "TV mode (full-screen)"
            screentip = GetText(g_arr_Text, "EXCELOPT012")
        Case "combo01InstalledRaces" 'Combobox "Race selection"
            screentip = GetText(g_arr_Text, "TIP024")
        Case "btn30GALOPPSIM" 'Button "Start race"
            screentip = GetText(g_arr_Text, "TIP025")
        Case "btn40GALOPPSIM" 'button "Info"
            screentip = GetText(g_arr_Text, "TIP026")
        Case "btn60GALOPPSIM" 'button "Play Movie"
            screentip = GetText(g_arr_Text, "BTN030")
    End Select
End Sub

'xxxxx
Private Sub AI_GetSupertip(control As IRibbonControl, ByRef supertip)
    Select Case control.id
        Case "cb01aGALOPPSIM" 'Checkbox "Excel mode"
            supertip = GetText(g_arr_Text, "TIP013")
        Case "cb01bGALOPPSIM" 'Checkbox "TV mode with Excel menu strip
            supertip = GetText(g_arr_Text, "TIP014")
        Case "cb01cGALOPPSIM" 'Checkbox "TV mode (full-screen)"
            supertip = GetText(g_arr_Text, "TIP015")
        Case "btn40GALOPPSIM" 'Button "Info"
            supertip = GetText(g_arr_Text, "TIP027")
        Case "btn60GALOPPSIM" 'Button "Movie"
            supertip = GetText(g_arr_Text, "BTN031")
    End Select
End Sub

'xxxxx Callbacks for button status
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
    End Select
End Sub

'xxxxx Initialwerte der Checkboxen im Menüband (getPressed)
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

'xxxxx Checkboxen im Menüband (onAction)
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

'xxxxx Ribbon button "Race options"
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

'xxxxx Startbutton im Menüband
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
            Call ShowNewRaceScreen
    End If

    Call NewRace
End Sub

'xxxxx Ergebnis-Button im Menüband
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

'xxxxx Ribbon button "Winner"
Private Sub AI_Winner(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call ShowWinnerPhoto
End Sub

'Show the photo of the winner
Public Sub ShowWinnerPhoto()
    If objRace.STARTED Then
        Call basAuxiliary.ActivateRaceSheet
        Call DrawWinnerPhoto
        Call basAuxiliary.Scroll(objRace.METRES + m_intLeftColumns + 16 + 160 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR), m_intTopRows + (objRace.NUMBER_ENROLLED * 2 + 9))
        If g_strPlayMode = "RS" Then frmRS_navigation.show (vbModeless)
    End If
End Sub

'Show the ranking list with the race results
Public Sub ShowRankingList()
    If objRace.STARTED Then
        Call basAuxiliary.ActivateRaceSheet
        Call FormatPhotoAndRanking
        Call basAuxiliary.Scroll(objRace.METRES + m_intLeftColumns + 15 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR), m_intTopRows + (objRace.NUMBER_ENROLLED * 2 + 9))
        Call DrawRankingList(False)
        If g_strPlayMode = "RS" Then frmRS_navigation.show (vbModeless)
    End If
End Sub

'xxxxx Zielfoto-Button im Menüband
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
Public Sub ShowFinishPhoto()
    If objRace.STARTED Then
        Call basAuxiliary.ActivateRaceSheet
        Call FormatPhotoAndRanking
        Call basAuxiliary.Scroll(objRace.METRES + m_intLeftColumns + 15 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR), m_intTopRows + 1)
        Call DrawFinishPhoto
        If g_strPlayMode = "RS" Then frmRS_navigation.show (vbModeless)
    End If
End Sub

'xxxxx Wett-Button im Menüband
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

'xxxxx Info-Button im Menüband
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

'xxxxx AI edition: Adapt Excel settings according to the selected mode
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

'xxxxx
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
'        .img_info_team02.ControlTipText = GetText(g_arr_Text, "TEAM006")
        'Paul
        .lbl_info_team03a.caption = GetText(g_arr_Text, "TEAM007")
        .lbl_info_team03b.caption = GetText(g_arr_Text, "TEAM008")
'        .img_info_team03.ControlTipText = GetText(g_arr_Text, "TEAM009")
        'Michael
        .lbl_info_team04a.caption = GetText(g_arr_Text, "TEAM010")
        .lbl_info_team04b.caption = GetText(g_arr_Text, "TEAM011")
'        .img_info_team04.ControlTipText = GetText(g_arr_Text, "TEAM012")
        'Meike
        .lbl_info_team05a.caption = GetText(g_arr_Text, "TEAM013")
        .lbl_info_team05b.caption = GetText(g_arr_Text, "TEAM014")
'        .img_info_team05.ControlTipText = GetText(g_arr_Text, "TEAM015")
        'Natalie
        .lbl_info_team06a.caption = GetText(g_arr_Text, "TEAM016")
        .lbl_info_team06b.caption = GetText(g_arr_Text, "TEAM017")
'        .img_info_team06.ControlTipText = GetText(g_arr_Text, "TEAM018")
        'Atanas
        .lbl_info_team07a.caption = GetText(g_arr_Text, "TEAM019")
        .lbl_info_team07b.caption = GetText(g_arr_Text, "TEAM020")
'        .img_info_team07.ControlTipText = GetText(g_arr_Text, "TEAM021")
        'Duncan
        .lbl_info_team08a.caption = GetText(g_arr_Text, "TEAM022")
        .lbl_info_team08b.caption = GetText(g_arr_Text, "TEAM023")
'        .img_info_team08.ControlTipText = GetText(g_arr_Text, "TEAM024")
        'Enno
        .lbl_info_team09a.caption = GetText(g_arr_Text, "TEAM025")
        .lbl_info_team09b.caption = GetText(g_arr_Text, "TEAM026")
'        .img_info_team09.ControlTipText = GetText(g_arr_Text, "TEAM027")
        'Farida
        .lbl_info_team10a.caption = GetText(g_arr_Text, "TEAM028")
        .lbl_info_team10b.caption = GetText(g_arr_Text, "TEAM029")
'        .img_info_team10.ControlTipText = GetText(g_arr_Text, "TEAM030")
        'Jan
        .lbl_info_team11a.caption = GetText(g_arr_Text, "TEAM031")
        .lbl_info_team11b.caption = GetText(g_arr_Text, "TEAM032")
'        .img_info_team11.ControlTipText = GetText(g_arr_Text, "TEAM033")
        
        'Vertical scrollbar
        With .multiPage_info.Pages(1)
            .ScrollBars = fmScrollBarsVertical
            .ScrollHeight = 610
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

'xxxxx Warnhinweis-Button im Menüband
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

'xxxxx Movie button (Ribbon)
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

'xxxxx
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

    'Deactivate screen updating
    Application.ScreenUpdating = False
    
    'Paint title picture
    k = basAuxiliary.GetColumn(g_wksPIC, "AI_TITLE")
    m = 2 'initial row for reading the picture
        
    For i = 1 To 40
        For j = 1 To 100
            g_wksRace.Cells(i, j).Interior.color _
                = g_wksPIC.Cells(m, k).Value
            m = m + 1 'next row on the worksheet "Pic"
        Next j
    Next i
        
    'Place the cursor far away (in the upper right corner of the screen)
    Call CursorAway

    'Activate screen updating
    Application.ScreenUpdating = True

    objRace.STARTED = False
        
End Sub

''Ad hoc race button
'Public Sub AI_adhocRace(control As IRibbonControl)
'    Call basAdhocRace.AdhocRace
'End Sub

'xxxxx Ende-Button im Menüband
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
    
    'Blatt löschen
    Call AI_DeleteWorksheets
    Unload frmBettingAnalysis
    
    'Reset Excel ribbon
    g_RibbonGaloppSim.Invalidate

    'Reset Excel options
    Call ResetExcelOptions
    Application.ScreenUpdating = True
End Sub

'xxxxx Language button "DE"
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

'xxxxx Language button "EN"
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

''xxxxx Language button "RU"
'Private Sub AI_LanguageRU(control As IRibbonControl)
'
'    'Check whether algorithms are allowed
'    If objOption.STOP_ALG Then
'        If basAuxiliary.AllowAlgorithms = False Then
'            objOption.STOP_ALG = False
'        Else
'            Exit Sub
'        End If
'    End If
'
'    objOption.LANGUAGE = "RU"
'    Call ChangeLanguage
'End Sub
'
''xxxxx Language button "CH"
'Private Sub AI_LanguageCH(control As IRibbonControl)
'
'    'Check whether algorithms are allowed
'    If objOption.STOP_ALG Then
'        If basAuxiliary.AllowAlgorithms = False Then
'            objOption.STOP_ALG = False
'        Else
'            Exit Sub
'        End If
'    End If
'
'    objOption.LANGUAGE = "CH"
'    Call ChangeLanguage
'End Sub
'
'xxxxx Language button "BG"
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
    End Select
End Sub

'xxxxx Anzahl der installierten Rennen auslesen
Private Sub AI_InstalledRaces_getItemCount(control As IRibbonControl, ByRef returnedVal)
    
    Dim wksR As Worksheet
    Dim cnt As Long
    
    Set g_colRacesInstalled = Nothing
    Set g_colRacesInstalled = New Collection
    
    For Each wksR In ThisWorkbook.Worksheets
        If left(wksR.name, 5) = "race_" Then cnt = cnt + 1
    Next wksR
    
    returnedVal = cnt
    
End Sub

'xxxxx Namen der installierten Rennen auslesen
Private Sub AI_InstalledRaces_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)

    Dim wksCheck As Worksheet
    Dim cnt As Long
    
    For Each wksCheck In ThisWorkbook.Worksheets
        If left(wksCheck.name, 5) = "race_" Then cnt = cnt + 1
        If cnt = index + 1 Then
            With wksCheck
                g_colRacesInstalled.Add .name
                returnedVal = .Cells(basAuxiliary.GetRow(wksCheck, "RACE NAME"), 2).Value & " " & _
                                .Cells(basAuxiliary.GetRow(wksCheck, "YEAR"), 2).Value & " (" & _
                                .Cells(basAuxiliary.GetRow(wksCheck, "DISTANCE METRES"), 2).Value & "m) - " & .Cells(basAuxiliary.GetRow(wksCheck, "TRACK LOCATION"), 2).Value
                End With
            Exit For
        End If
    Next wksCheck

End Sub

'xxxxx Set default race
Private Sub AI_InstalledRaces_GetSelectedItemID(control As IRibbonControl, ByRef itemID As Variant)
    If objRace.SELECTED = "" Then
        objRace.SELECTED = g_colRacesInstalled(1) 'take the first race of the collection
    End If
    itemID = objRace.SELECTED
End Sub

'xxxxx Ausgewähltes Rennen setzen
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
        ThisWorkbook.Saved = True 'Do not save the workbook!
        'https://support.microsoft.com/en-us/help/213428/how-to-suppress-save-changes-prompt-when-you-close-a-workbook-in-excel
    End If
End Sub
