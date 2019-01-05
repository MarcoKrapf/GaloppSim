Attribute VB_Name = "basMainCode"
Option Explicit
Option Private Module

'This module contains the main code of the race simulator with most of the logic

'GALOPPSIM - Version 149.10 (January 2019)
'Horse racing simulator for Microsoft Excel
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
        
'CONSTANTS AND VARIABLES USED FOR DEVELOPMENT
Public Const DEVMODE As Boolean = False 'Set true to skip waiting commands (Application.Wait)

'GLOBAL CONSTANTS AND VARIABLES
'------------------------------
Public Const g_c_tool As String = "GaloppSim" 'Name of the tool
Public Const g_c_version As String = "(v149.10)" 'Version of the tool
Public Const g_c_email As String = "info@galoppsim.racing" 'Contact e-mail address
Public Const g_c_defaultRaceOptionsFile As String = "RACEOPTIONS" 'File name for the race options
Public Const g_c_defaultFileType As String = ".GALOPPSIM" 'File type for GaloppSim files

Public objRace As clsRace 'Object with all the race data
Public objOption As clsOptions 'Object with all the race and excel options
Public g_RibbonGaloppSim As IRibbonUI 'Custom ribbon
Public g_strPlayMode As String ' "AI" = AddIn (.xlam) / "RS" = Run Simple (.xlsm)
Public g_defaultPath As String 'Path for the Galoppsim files
Public g_objLabel As MSForms.Label 'Label used for different purposes
Public g_colRSbuttons As Collection 'Menu buttons in the RS edition
Public g_oleComboRaces As OLEObject 'ComboBox with installed races in the RS edition
Public g_arr_varHorses() As Variant 'All information about the horses
Public g_colBetSlips As Collection 'List with all betting slips
Public g_arr_Text() As Variant 'All text components
Public g_arr_Grammar(1 To 8) As String 'All animal grammar components
Public g_strMsgCaption As String 'Caption for a pop-up
Public g_strMsgText As String 'Text for a pop-up
Public g_strMsgButtons As String 'Buttons for a pop-up
Public g_strPlayerName As String 'Name of the player who places a bet
Public g_strButtonPressed As String 'Return value of the pressed button

'Existing worksheets with data
Public g_wksTEXT As Worksheet 'Worksheet with text components
Public g_wksPIC As Worksheet 'Worksheet with picture data

'Worksheets created at runtime
Public g_wksRace As Worksheet 'Worksheet for the race
Public g_wksMovie As Worksheet 'Worksheet for the movie
Public g_wksRaceData As Worksheet 'Worksheet with the race data

    
'VARIABLES AND CONSTANTS ON MODULE-LEVEL
'---------------------------------------

Dim m_wksTec As Worksheet 'Worksheet with technical data (speed, tactics...)
Dim m_wksCheck As Worksheet 'Variable to check whether the table sheet GALOPPSIM exists
Dim m_arr_varPhotofinish() As Variant 'Position of each horse for the finish photo
Dim m_arr_varResultsCalc() As Variant 'Calculation of the position at the finish line
Dim m_arr_varResults() As Variant 'Result list
Dim m_colRacesInstalled As Collection 'List of installed races
Dim m_arr_varAdv() As Variant 'Advertising sequence
Dim m_intTopRows As Integer 'Number of rows at the top of the worksheet (used for the menu in the RS edition)
Dim m_intLeftColumns As Integer 'Number of columns left of the start boxes
Dim m_intColumsAfterFinish As Integer 'Number of columns behind the finish line
Dim m_intTrackCellHeight As Integer 'Cell height (race track)
Dim m_dblTrackCellWidth As Double 'Cell width (race track)
Dim m_dblRankingsWidth As Double 'Cell width of the finish photo and the ranking list
Dim m_intAdvertisingHeight As Integer 'Row height of the advertising area
Dim m_intFontSize As Integer 'Font size of the horse names and the hoof prints
Dim m_lngSpeedCondHigh As Long, m_lngSpeedCondLow As Long 'Variables for the speed range of the daily form of the horses
Dim m_lngSpeedLoopHigh As Long, m_lngSpeedLoopLow As Long 'Variables for the range of the randomly assigned speed per step
Dim m_lngSpeedTacticsHigh As Long, m_lngSpeedTacticsMedium As Long, m_lngSpeedTacticsLow As Long 'Variables for the speed range of each phase if racing tactics are active
Dim m_intHorsesStarting As Integer 'Number of horses that start
Dim m_intHorsesRunning As Integer 'Number of horses currently running
Dim m_byteFavourite(1 To 3) As Byte 'Array for three predicted favourites of the race
Dim m_dblFavCalc(1 To 3) As Double 'Array for calculating the favourites
Dim m_intHorsesFinishing As Integer 'Variable that counts how many horses arrive at the finish line in one loop
Dim m_intFinishLoop As Integer 'Variable that counts the finishing loop in which a placement was calculated
Dim m_intPlace As Integer 'Placement in the finish
Dim m_strWinner As String 'Name(s) of the horse(s) in 1st place
Dim m_blnDeadHeat As Boolean 'Dead heat (more than one horse has won)
Dim m_blnWin As Boolean 'Flag indicating that a horse has won the race
Dim m_blnPhotofinish As Boolean 'Flag indicating whether there is a photo finish
Dim m_blnExactFinish As Boolean 'Variable for the exact calculation at the finish line 'ToDo... NRA
Dim m_intPosLeader As Integer 'Position of the leading horse
Dim m_strNameLeader As String 'Name of the leading horse
Dim i As Integer, j As Integer, k As Integer, m As Integer 'Counting variables for loops
Dim z As Long 'Auxiliary variable for loops


'PROCEDURES
'----------

'Starting procedure (RS edition)
Public Sub RS_NewRace()
    ThisWorkbook.Worksheets("RS").Activate 'Show the launch screen
                                            'while preparing all elements
    m_intTopRows = 8 'Space on the top of the worksheet for the control elements
                        'like the dropdown and menu buttons
    Call AssignDataSheets 'Assign the worksheets "TEXT", "PIC" and "TEC"
    Call GetTextComponents 'Get texts according to the selected language
    Call CreateRaceSheet 'Create the worksheet "GALOPPSIM"
    Call RS_StartScreen 'Create the start screen with the control elements
    Call RS_AddCommandButtons 'Add menu buttons
    Call RS_AddComboBox 'Add dropdown for the race selection
End Sub

'Starting procedure for a new race
Private Sub NewRace()

    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    'Close pop-ups if visible
    If frmBettingAnalysis.Visible Then Unload frmBettingAnalysis 'Betting analysis
    If frmRS_navigation.Visible Then Unload frmRS_navigation 'Navigation panel (RS edition only)
    
    If g_strPlayMode = "AI" Then
        Call AssignDataSheets 'Assign worksheets with texts, pictures and technical data
        Call GetTextComponents 'Get texts according to the selected language
    End If
    
    'Reset the betting slip collection
    Set g_colBetSlips = Nothing
    Set g_colBetSlips = New Collection

    Call GetRaceData 'Get the race data from the worksheet with the selected race
    Call AnimalGrammar 'Get grammar components according to the selected language
    Call AssignBasicValues 'Get basic values
    Call GetHorseData 'Get data about the horses according to the selected race
    Call CheckBettingAllowed 'Check whether betting is allowed for this race
    Call UserFormSTART 'Show Start-UserForm
        
    If objRace.STARTED Then
        If g_strPlayMode = "AI" Then
            Call CreateRaceSheet 'Tabellenblatt "GALOPPSIM"
            Call AI_ExcelModeStart
            Call CursorAway 'Place the cursor far away (in the upper right corner of the screen)
        End If
        
        If g_strPlayMode = "RS" Then
            Cells.Clear 'Clear the whole worksheet
            With g_wksRace.Cells(2, 2) 'Write title
                .Font.name = "Arial Black"
                .Value = g_c_tool & " " & g_c_version
            End With
            Call RS_HideNavigationPanel 'hide the navigation area
        End If

        Call DrawRaceTrack 'Geläuf zeichnen
        Call DrawHorseNames 'Pferdenamen am Start und im Ziel wenn angehakt
        If objOption.RACE_INFO And objOption.RACE_INFO_POP Then Call RaceInfoPopup 'Show pop-up with race info if checked
        Call RaceWelcome 'Popup zu Rennbeginn
        Call StartingGrid 'Pferde in Boxen stellen
        Call RacePresentation 'Pferde vorstellen
        Call RunRace 'Rennstart
        Call NotFinished 'Find the horses that did not finish
        Call RaceFinished 'Info pop-up after the race
        Call ShowRankingList(True) 'Ergebnistafel
        Call DrawWinnerPhoto 'Grafik
        If objOption.BET_PLACED And objOption.BET_ANALYSIS Then Call UserFormAnalyseBetSlips 'Analyse bet slips
        
        If g_strPlayMode = "RS" Then 'Show the navigation area when running RS mode
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
            Call RS_ShowNavigationPanel(True)
        End If

        If g_strPlayMode = "AI" Then
            g_RibbonGaloppSim.Invalidate 'refresh the status buttons
            Call AI_ExcelModeEnd
        End If
        
    End If

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Get the selected race from the dropdown
Private Sub GetRace()
    If g_strPlayMode = "RS" Then
        objRace.SELECTED = g_oleComboRaces.Object.Value
    Else
        '(currently not used for AI edition)
    End If
End Sub

'Procedure for painting a horse
'ToDo... ausführlich kommentieren
Private Sub PaintHorse(ByVal row As Integer, tail As Integer, color As Variant)
    Dim horseColor As Long
    Dim check_array As Boolean
    check_array = False
    If IsArray(color) Then
        If UBound(color) > 6 Then
            Dim count As Integer, countInitial As Integer
            Dim currentColor As Long
            count = 0
            Do While count < 8
                countInitial = count
                currentColor = color(count)
                Do
                    If count = 7 Then Exit Do
                    If currentColor = color(count + 1) Then count = count + 1 Else Exit Do
                Loop
                If count = countInitial Then
                    Cells(row, tail + count).Interior.color = currentColor
                Else
                    Range(Cells(row, tail + countInitial), Cells(row, tail + count)).Interior.color = currentColor
                End If
                count = count + 1
            Loop
            check_array = True
        End If
    End If
    
    If Not check_array Then
        If IsNumeric(color) Then
            horseColor = color
        Else    'no valid color has been provided, use brown instead:
            horseColor = 3291720
        End If
            Range(Cells(row, tail), Cells(row, tail + 7)).Interior.color = horseColor
    End If
End Sub

Private Sub CheckBettingAllowed()
    If objOption.BET_MODE = True And objRace.BETTING_ALLOWED = "N" Then
    'Show info pop-up
        'Set the button mode
        g_strMsgButtons = "OK"
        'Assign the text for the pop-up
        g_strMsgCaption = objRace.RACE_NAME & " " & objRace.RACE_YEAR
        g_strMsgText = " " & UCase(GetText(g_arr_Text, "BET050")) & "."
        'Display the pop-up
        frmMsg_Info.show (vbModal) 'modal
    End If
End Sub

Public Sub RS_AddComboBox()
    'Add combobox to the navigation panel
    Call RS_AddComboboxRaces("CBraces", 196, 15, 597, 22) '"name(ID)", left, top, width, height
End Sub

'Add controls in RS mode
Public Sub RS_AddCommandButtons()
    'Error handling
    On Error GoTo ERRORHANDLING 'REIN

    Dim captionStart As String
    captionStart = basAuxiliary.getCaptionStartBtn(objOption.BET_MODE)
    
    'Add buttons to the navigation panel
        Set g_colRSbuttons = New Collection
        
        '"name(ID)", left, top, width, height, font-size, font:bold, _
            background-color (hex), text(xxx is the initial caption)
        Call RS_AddButton("raceoptions", 15, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN001"))
        Call RS_AddButton("exceloptions", 99, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN002"))
        Call RS_AddButton("startrace", 196, 40, 81, 49, 11, True, 52377, GetText(g_arr_Text, captionStart)) 'green button
        Call RS_AddButton("finishphoto", 280, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN004"))
        Call RS_AddButton("results", 364, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN005"))
        Call RS_AddButton("winner", 448, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN006"))
        Call RS_AddButton("bets", 532, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN007"))
        Call RS_AddButton("language", 629, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "LANGUAGE001"))
        Call RS_AddButton("info", 713, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN009"))
        Call RS_AddButton("warning", 797, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN010"))
        Call RS_AddButton("movie2017", 881, 40, 81, 49, 11, False, &HFFFFFF, GetText(g_arr_Text, "BTN011"))

        Call RS_InactivateCommandButtons 'inactivate some buttons
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

Private Sub RS_InactivateCommandButtons()
    'Set buttons inactive
        With g_wksRace
            .OLEObjects("finishphoto").Object.Enabled = False
            .OLEObjects("results").Object.Enabled = False
            .OLEObjects("winner").Object.Enabled = False
            .OLEObjects("bets").Object.Enabled = False
        End With
End Sub

'Click on a button in the Run Simple Edition
Public Sub RS_ExecuteClick(name As String)
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Select Case name

        Case "startrace"
            'Leave the current race?
            If objRace.STARTED Then
                
'                    'A MessageBox cannot handle unicode, so for example Cyrillic characters are displayed as question marks
'                    If MsgBox((GetText(g_arr_Text, "RACE003") & ": " & g_wksRace.OLEObjects("CBraces").Object.text), _
'                        vbOKCancel, g_c_TOOL) = vbCancel Then Exit Sub
                    
                'Set the button mode
                g_strMsgButtons = "CancelOK"
                'Assign the text for the pop-up
                g_strMsgCaption = g_c_tool
                g_strMsgText = GetText(g_arr_Text, "RACE003") & ": " & g_wksRace.OLEObjects("CBraces").Object.text
                'Display the pop-up
                frmMsg_MultiPurpose.show (vbModal) 'modal
                'Evaluate the return value
                If g_strButtonPressed = "CANCEL" Then Exit Sub
                
                Call ShowNewRaceScreen
            End If
            
            Call GetRace
            Call GetRaceData
            
            Call RS_InactivateCommandButtons 'inactivate some buttons
            Call NewRace
        
        Case "finishphoto"
            Call ShowFinishPhoto
        Case "results"
            Call RankingList
        Case "winner"
            Call ShowWinnerPhoto
        Case "bets"
            Call ShowBets
        Case "raceoptions"
            Call GetRace
            Call AnimalGrammar
            frmOptionsRace.show (vbModal) 'display UserForm (modal)
        Case "exceloptions"
            frmOptionsExcel.show (vbModal) 'display UserForm (modal)
        Case "language"
            frmRS_languages.show (vbModal) 'display UserForm (modal)
            If objRace.SELECTED = "" Then
                objRace.SELECTED = g_oleComboRaces.Object.Value
                Call GetRaceData
                Call AnimalGrammar
            End If
            Call ChangeLanguage
        Case "info"
            Call ShowInfo
        Case "warning"
            Call warning
        Case "movie2017"
            Call GaloppSimMovie2017
    End Select
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

Private Sub ShowNewRaceScreen()
    With g_wksRace.UsedRange
        .ColumnWidth = ZoomLevelPictures()(0)
        .RowHeight = ZoomLevelPictures()(1)
        .Clear
    End With
    If g_strPlayMode = "RS" Then Call RS_HideNavigationPanel 'hide the navigation area
    If g_strPlayMode = "AI" Then Call AI_ExcelModeStart
    ActiveWindow.ScrollColumn = 1 'scroll to the left
    Call PaintPicture(g_wksPIC, "NEWRACE", 100, 40, 1, 1)  'paint title picture
    'Place the cursor far away (in the upper right corner of the screen)
        Call CursorAway
End Sub

Private Sub PaintPicture(wks As Worksheet, ByVal pic As String, cols As Integer, rows As Integer, top As Integer, left As Integer)

    'Deactivate screen updating
        Application.ScreenUpdating = False
        
    'Paint a picture (get data from the worksheet "PIC")
        k = basAuxiliary.GetColumn(wks, pic)
        
        m = 2 'Initial row for reading the picture
        
        For i = top To rows
            For j = left To cols
                With g_wksRace.Cells(i, j)
                    .Clear
                    .Interior.color = wks.Cells(m, k).Value
                End With
                m = m + 1 'Next row on the worksheet "PIC"
            Next j
        Next i

    'Activate screen updating
        Application.ScreenUpdating = True
        
End Sub

'Add new RS button
Private Sub RS_AddButton(n As String, l As Integer, t As Integer, w As Integer, _
                            h As Integer, fs As Integer, fb As Boolean, bc As Long, c As String)
    Dim oleRSbutton As OLEObject
    Dim objRSbutton As clsRSbutton
    
    Set oleRSbutton = g_wksRace.OLEObjects.Add(classtype:="Forms.CommandButton.1", _
    left:=l, top:=t, Width:=w, Height:=h)
    
    'Assign properties to the button using nested "With" commands
    With oleRSbutton
        .name = n 'ID
        With .Object
            .Caption = c '[ Full command: oleRSbutton.Object.Caption = c ]
            .Font.Size = fs '[ Full command: oleRSbutton.Object.Font.Size = fs ]
            .Font.Bold = fb
            .BackColor = bc
            .WordWrap = True
            .TakeFocusOnClick = False
        End With
        .Placement = xlFreeFloating '[ Full command: oleRSbutton.Placement = xlFreeFloating ]
        .Visible = True
    End With
    
    Set objRSbutton = New clsRSbutton
    Set objRSbutton.RSButtonObject = oleRSbutton.Object
    objRSbutton.RSbtnID = n

    g_colRSbuttons.Add objRSbutton
 
End Sub

'Add new RS combobox with races
Private Sub RS_AddComboboxRaces(n As String, l As Integer, t As Integer, w As Integer, h As Integer)

    Dim wksCheck As Worksheet
    
    Set m_colRacesInstalled = Nothing
    Set m_colRacesInstalled = New Collection

    Set g_oleComboRaces = g_wksRace.OLEObjects.Add(classtype:="Forms.ComboBox.1", _
        left:=l, top:=t, Width:=w, Height:=h)
    
    With g_oleComboRaces
        .name = n 'ID
        .Placement = xlFreeFloating
        .Object.ColumnCount = 2 'column 0: g_wksRace-name // column 1: visible name
        .Object.ColumnWidths = "0 Pt" 'width of the column with the race name --> hidden
        .Object.Style = fmStyleDropDownList 'Allow only values from the item list, no free entries
        .Visible = True
    End With
    
    'Populate the dropdown with installed races
    For Each wksCheck In ThisWorkbook.Worksheets
        If left(wksCheck.name, 5) = "race_" Then
            m_colRacesInstalled.Add wksCheck.name
            With g_oleComboRaces.Object
            .AddItem
            .List(.ListCount - 1, 0) = wksCheck.name
            .List(.ListCount - 1, 1) = wksCheck.Cells(basAuxiliary.GetRow(wksCheck, "RACE NAME"), 2).Value & " " & _
                                    wksCheck.Cells(basAuxiliary.GetRow(wksCheck, "YEAR"), 2).Value & " (" & _
                                    wksCheck.Cells(basAuxiliary.GetRow(wksCheck, "DISTANCE METERS"), 2).Value & "m) - " & wksCheck.Cells(basAuxiliary.GetRow(wksCheck, "TRACK LOCATION"), 2).Value
            End With
        End If
    Next wksCheck
    
    'Set default race
    g_oleComboRaces.Object.ListIndex = 0 'take the first race

End Sub

Private Sub RS_HideNavigationPanel()
    Dim oleObj As OLEObject

    For Each oleObj In g_wksRace.OLEObjects
        oleObj.Visible = False
    Next oleObj
    
    g_wksRace.Range(rows(1), rows(m_intTopRows)).Hidden = True
End Sub

Public Sub RS_ShowNavigationPanel(popup As Boolean)
    Dim oleObj As OLEObject

    g_wksRace.Range(rows(1), rows(m_intTopRows)).Hidden = False

    For Each oleObj In g_wksRace.OLEObjects
'        Debug.Print "show... " & oleObj.name
        oleObj.Visible = True
    Next oleObj

    If popup = True Then frmRS_navigation.show (vbModeless) 'modeless

End Sub

'Start screen for the RS edition
Private Sub RS_StartScreen()
    'Paint title picture
        Call PaintPicture(g_wksPIC, "RUNSIMPLE", 100, 40, 1, 1)
    'Set the column width and row height dependent on the screen size
        With g_wksRace.Range(Columns(1), Columns(100))
            .ColumnWidth = ZoomLevelPictures()(0)
            .RowHeight = ZoomLevelPictures()(1)
            .rows(1).RowHeight = 8 'height of the top line
            With .rows(2)
                .Font.name = "Arial Black" 'for getting the dropdown with the races bold
                .EntireRow.AutoFit
            End With
        End With
    'Write title
        g_wksRace.Cells(2, 2).Value = g_c_tool & " " & g_c_version
    'Place the cursor far away (in the upper right corner of the screen)
        Call CursorAway
End Sub

'Create worksheet for the race
Private Sub CreateRaceSheet()
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN

    'Prüfen ob es schon ein Tabellenblatt gibt
    For Each m_wksCheck In ActiveWorkbook.Worksheets
        If m_wksCheck.name = "GALOPPSIM" Then 'Tabellenblatt ist schon da
            Application.DisplayAlerts = False 'Warnmeldungen ausschalten
            m_wksCheck.Delete 'Tabellenblatt löschen
            Application.DisplayAlerts = True 'Warnmeldungen einschalten
        End If
    Next m_wksCheck
    'Neues Tabellenblatt generieren
        Set g_wksRace = ActiveWorkbook.Worksheets.Add(Before:=Sheets(1)) '(after:=Sheets(Sheets.Count))
        g_wksRace.name = "GALOPPSIM"
        g_wksRace.Activate

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Assign the worhsheets of this workbook
Private Sub AssignDataSheets()
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    Set g_wksTEXT = tableTEXT
    Set g_wksPIC = tablePIC
    Set m_wksTec = ThisWorkbook.Worksheets("TEC")
    'Alternative but less stable assignment:
'    Set m_wksTec = ThisWorkbook.Worksheets("TEC")
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Grunddaten einlesen
Private Sub AssignBasicValues()
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
        
    'Determine the zoom level
    If objOption.ZOOM_LEVEL = 0 Then
        objOption.ZOOM_LEVEL = ZoomLevelRecommendation()
    Else
        Dim byteOpt As Byte
        byteOpt = ZoomLevelRecommendation() 'get the perfect level by calling the function

        If objOption.ZOOM_LEVEL <> byteOpt Then 'compare the selected value with the recommendation
                    'Set the button mode
                    g_strMsgButtons = "YesNo"
                    'Assign the text for the pop-up
                    g_strMsgCaption = objRace.RACE_NAME & " " & objRace.RACE_YEAR
                    g_strMsgText = GetText(g_arr_Text, "ZOOM007") & vbNewLine & vbNewLine & _
                                    GetText(g_arr_Text, "ZOOM008") & ": " & ZoomLevelText(objOption.ZOOM_LEVEL) & vbNewLine & _
                                    GetText(g_arr_Text, "ZOOM009") & ": " & ZoomLevelText(byteOpt) & vbNewLine & vbNewLine & _
                                    GetText(g_arr_Text, "ZOOM010")
                    'Display the pop-up
                    frmMsg_MultiPurpose.show (vbModal) 'modal
                    'Evaluate the return value
                    If g_strButtonPressed = "YES" Then objOption.ZOOM_LEVEL = byteOpt 'adapt the value
        End If
    End If
    
    'Assign the values for the zoom level
        Select Case objOption.ZOOM_LEVEL
            Case 1
                m_intTrackCellHeight = 6 'cell width of the race track
                m_dblTrackCellWidth = 0.2 'cell length of the race track
                m_intFontSize = 5 'font size of the horse names and the hoof prints
                m_dblRankingsWidth = 3 'width of the finish photo and the ranking list
                m_intAdvertisingHeight = 2 'row height of the advertising area
            Case 2
                m_intTrackCellHeight = 9 'cell width of the race track
                m_dblTrackCellWidth = 0.3 'cell length of the race track
                m_intFontSize = 8 'font size of the horse names and the hoof prints
                m_dblRankingsWidth = 4.5 'width of the finish photo and the ranking list
                m_intAdvertisingHeight = 5 'row height of the advertising area
            Case 3
                m_intTrackCellHeight = 12 'cell width of the race track
                m_dblTrackCellWidth = 0.4 'cell length of the race track
                m_intFontSize = 10 'font size of the horse names and the hoof prints
                m_dblRankingsWidth = 6 'width of the finish photo and the ranking list
                m_intAdvertisingHeight = 6 'row height of the advertising area
        End Select
        
    'Geschwindigkeitsspanne des kursfristigen Faktors pro Schleifendurchlauf
        m_lngSpeedLoopHigh = m_wksTec.Range("A2").Value 'Richtwert 50300
        m_lngSpeedLoopLow = m_wksTec.Range("A3").Value 'Richtwert 49700
        
    'Geschwindigkeitsspanne der Form
        m_lngSpeedCondHigh = m_wksTec.Range("B2").Value 'Richtwert 50010
        m_lngSpeedCondLow = m_wksTec.Range("B3").Value 'Richtwert 49990
    
    'Geschwindigkeiten in den Rennphasen (Taktik)
        m_lngSpeedTacticsHigh = m_wksTec.Range("C2").Value 'Richtwert 50050
        m_lngSpeedTacticsMedium = m_wksTec.Range("C3").Value '50000
        m_lngSpeedTacticsLow = m_wksTec.Range("C4").Value 'Richtwert 49950
    
    'Spalten bis zu den Boxen (mind. 7, Richtwert 10)
        m_intLeftColumns = 11
    'Spalten hinter der Zielline (mind. 7, je nach Auslauf
        m_intColumsAfterFinish = 5
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Grunddaten über das Rennen einlesen
Private Sub GetRaceData()
    
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    'Tabellenblatt mit ausgewähltem Rennen zuweisen
    Set g_wksRaceData = ThisWorkbook.Worksheets(objRace.SELECTED)
    
    'Grunddaten aus Tabellenblatt einlesen
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
        objRace.METERS = .Cells(basAuxiliary.GetRow(g_wksRaceData, "DISTANCE METERS"), 2).Value 'Race distance
        objRace.STARTING_GATE = .Cells(basAuxiliary.GetRow(g_wksRaceData, "STARTING GATE"), 2).Value 'Starting gate (yes or no)
        objRace.NUMBER_ENROLLED = .Cells(basAuxiliary.GetRow(g_wksRaceData, "NUMBER OF STARTERS"), 2).Value 'Number of horses
        objRace.LANES_FIX_OR_RANDOM = .Cells(basAuxiliary.GetRow(g_wksRaceData, "LANES FIX OR RANDOM"), 2).Value 'Lanes fix or random
        objRace.COLOURS_FIX_OR_RANDOM = .Cells(basAuxiliary.GetRow(g_wksRaceData, "COLOURS FIX OR RANDOM"), 2).Value 'Horse colours fix or random
        objRace.ODDS_FIX_FIX_OR_RANDOM = .Cells(basAuxiliary.GetRow(g_wksRaceData, "ODDS FIX OR RANDOM"), 2).Value 'Odds fix or random
        objRace.ADVERTISING = .Cells(basAuxiliary.GetRow(g_wksRaceData, "ADVERTISING"), 2).Value 'Advertising (yes or no)
        objRace.BETTING_ALLOWED = .Cells(basAuxiliary.GetRow(g_wksRaceData, "BETTING ALLOWED"), 2).Value 'Betting allowed (yes or no)
    End With
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Advertisement data
Private Sub GetAdvertisementData()

    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN

    Dim col As Integer 'Column with the advertisement data
    col = basAuxiliary.GetColumn(g_wksRaceData, "ADVERTISEMENT")

        j = g_wksRaceData.Cells(rows.count, col).End(xlUp).row - 1 'Last row oft the column ADVERTISEMENT
        ReDim m_arr_varAdv(1 To j) 'Location of the advertisement data
        For i = 1 To j
            For k = 1 To g_wksPIC.Cells(1, Columns.count).End(xlToLeft).column
                If g_wksRaceData.Cells(i + 1, col).Value = g_wksPIC.Cells(1, k).Value Then
                    m_arr_varAdv(i) = k 'Assign column number
                    Exit For
                End If
            Next k
        Next i
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Get the country name dependent on the selected race and language
Private Function GetCountryName(code As String, language As String) As String
    Dim col As Integer, row As Integer
    
    'Find the column on the worksheet
        col = basAuxiliary.GetColumn(g_wksTEXT, objOption.language) 'get the column with the language
        
    'Find the row on the worksheet
        row = basAuxiliary.GetRow(g_wksTEXT, code)
        
    'Return the country name
        If g_wksTEXT.Cells(row, col).Value = "" Then
            GetCountryName = g_wksTEXT.Cells(row, 5).Value 'return the English name if no name was found
        Else
            GetCountryName = g_wksTEXT.Cells(row, col).Value 'return the name according to the selected language
        End If
End Function

'Get the text components according to the selected language
Private Sub GetTextComponents()
    Dim col As Integer
    
    'Read the text components into arrays
        ReDim g_arr_Text(0 To 1, 1 To 2000) 'resize the array
        col = basAuxiliary.GetColumn(g_wksTEXT, objOption.language) 'get the column with the language
        For i = 1 To UBound(g_arr_Text, 2)
            g_arr_Text(0, i) = g_wksTEXT.Cells(i, 1).Value 'ID
            g_arr_Text(1, i) = g_wksTEXT.Cells(i, col).Value 'value
        Next i
    
End Sub

'Daten über die Pferde
Private Sub GetHorseData()
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    'Anzahl der Starter aus Tabellenblatt auslesen
        m_intHorsesStarting = Application.WorksheetFunction.CountIf(g_wksRaceData.Columns(6), "START")
    'Arrays anlegen
    ReDim g_arr_varHorses(1 To objRace.NUMBER_ENROLLED, 0 To 24) 'Alle Daten der Pferde
    ReDim m_arr_varPhotofinish(1 To objRace.NUMBER_ENROLLED, 0 To 4) 'Snapshot für Zielfoto/Fotofinish
    ReDim m_arr_varResults(0 To m_intHorsesStarting, 0 To 8) 'Ergebnisliste
    
    'In case of random lanes
        If objRace.LANES_FIX_OR_RANDOM = "R" Then
            Dim boxNr As Integer
            Dim inBox As Boolean
            Dim boxenNr() As Integer
            ReDim boxenNr(1 To objRace.NUMBER_ENROLLED)
            For i = 1 To objRace.NUMBER_ENROLLED
                boxenNr(i) = i
            Next i
        End If
    
    For i = 1 To objRace.NUMBER_ENROLLED
        Dim arr_color(0 To 7)
        Dim color As Integer
        Dim same_color As Boolean
        same_color = True
        g_arr_varHorses(i, 0) = g_wksRaceData.Cells(1 + i, 6).Value 'Status (START, CANCELLED, REFUSED...)
        g_arr_varHorses(i, 11) = g_wksRaceData.Cells(1 + i, 5).Value 'Startnummer
        g_arr_varHorses(i, 1) = g_wksRaceData.Cells(1 + i, 7).Value 'Name des Pferds
        If objRace.COLOURS_FIX_OR_RANDOM = "F" Then 'fix
            For color = 0 To 7
                arr_color(color) = g_wksRaceData.Cells(1 + i, 8 + color) 'Horse color
                If color > 0 Then
                    If Not g_wksRaceData.Cells(1 + i, 8 + color) = g_wksRaceData.Cells(1 + i, 7 + color) Then
                        same_color = False
                    End If
                End If
            Next color
            If same_color Then g_arr_varHorses(i, 2) = arr_color(0) Else g_arr_varHorses(i, 2) = arr_color
        Else 'random
            If g_arr_varHorses(i, 1) = "Loppsi" Then
                g_arr_varHorses(i, 2) = 192 'Loppsi is always red
            Else
                Randomize 'Zufallsgenerator zurücksetzen
                Dim random_color As Long
                Do
                    random_color = CLng(((10 - 1 + 1) * Rnd + 1) * 1000000)
                Loop Until random_color >= 0 And random_color <= 16777215 'allowed value range
                g_arr_varHorses(i, 2) = random_color 'apply randomly generated color for whole horse
            End If
        End If
        
        If objRace.LANES_FIX_OR_RANDOM = "R" Then 'lanes random
            inBox = False
            Do Until inBox = True
                Randomize
                boxNr = (Int((objRace.NUMBER_ENROLLED - 1 + 1) * Rnd + 1)) 'Zufallszahl
                If boxenNr(boxNr) <> 0 Then
                    g_arr_varHorses(i, 15) = boxNr 'Box aus der das Pferd startet
                    boxenNr(boxNr) = 0
                    inBox = True
                End If
            Loop
        Else 'lanes fix
            g_arr_varHorses(i, 15) = g_wksRaceData.Cells(1 + i, 4).Value 'Box aus der das Pferd startet
        End If
        
        g_arr_varHorses(i, 3) = m_intTopRows + 5 + 2 * g_arr_varHorses(i, 15) 'Zeilennummer auf der das Pferd läuft
        
        g_arr_varHorses(i, 4) = m_intLeftColumns + 12 'Fixe Startposition in der Box
        g_arr_varHorses(i, 5) = g_wksRaceData.Cells(1 + i, 16).Value 'Grundgeschwindigkeit des Pferds
                                                        '(linear von 1,50010 bis 1,49988)
        'Form des Pferds durch Zufallszahl festlegen
            Randomize 'Zufallsgenerator zurücksetzen
            g_arr_varHorses(i, 6) = (Int((m_lngSpeedCondHigh - m_lngSpeedCondLow + 1) * Rnd + m_lngSpeedCondLow) + 100000) / 100000 'Zufallszahl
        'Wettquote festlegen
            If objRace.ODDS_FIX_FIX_OR_RANDOM = "F" Then 'fix
                g_arr_varHorses(i, 17) = g_wksRaceData.Cells(1 + i, 17).Value
            Else 'random
                Randomize 'Zufallsgenerator zurücksetzen
                Do
                    g_arr_varHorses(i, 17) = Round(CInt((1 + (((Int((4 - 0 + 1) * Rnd + 0)) - 2) / 10)) _
                        * (150012 - (g_arr_varHorses(i, 6) * 100000)) * 10) / 5) * 5 'complex formula...
                Loop Until g_arr_varHorses(i, 17) >= 15 'minimum value
            End If
        'Schätzfehler für Balkenanzeige bei Wetten (+/-50)
            Randomize 'Zufallsgenerator zurücksetzen
            g_arr_varHorses(i, 18) = (Int((100 - 0 + 1) * Rnd + 0)) - 50 'random number between -50 and +50
        'Windschatten auf null setzen
            g_arr_varHorses(i, 22) = 0
        'Siegerbild
            g_arr_varHorses(i, 23) = g_wksRaceData.Cells(1 + i, 18).Value
        'Rennverhalten
            g_arr_varHorses(i, 24) = g_wksRaceData.Cells(1 + i, 19).Value
    Next i
    
    'Favoriten errmitteln aus Grundgeschwindigkeit und Form
        'Clear the entire array
            Erase m_dblFavCalc
            
'            'Alternatively: Clear the array fields one by one
'            m_dblFavCalc(1) = 0
'            m_dblFavCalc(2) = 0
'            m_dblFavCalc(3) = 0

'            'Alternatively: Clear the array fields using a loop
'            For i = 1 To 3
'                 m_dblFavCalc(i) = 0
'            Next i
            
        'Berechnung der drei Favoriten
            For i = 1 To objRace.NUMBER_ENROLLED
                If g_arr_varHorses(i, 0) = "START" Then
                    If g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) > m_dblFavCalc(1) Then
                        m_dblFavCalc(3) = m_dblFavCalc(2)
                        m_dblFavCalc(2) = m_dblFavCalc(1)
                        m_dblFavCalc(1) = g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6)
                        m_byteFavourite(3) = m_byteFavourite(2)
                        m_byteFavourite(2) = m_byteFavourite(1)
                        m_byteFavourite(1) = i 'Nummer des Favoriten
                    ElseIf g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) > m_dblFavCalc(2) Then
                        m_dblFavCalc(3) = m_dblFavCalc(2)
                        m_dblFavCalc(2) = g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6)
                        m_byteFavourite(3) = m_byteFavourite(2)
                        m_byteFavourite(2) = i 'Nummer eines weiteren Favoriten
                    ElseIf g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) > m_dblFavCalc(3) Then
                        m_dblFavCalc(3) = g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6)
                        m_byteFavourite(3) = i 'Nummer eines weiteren Favoriten
                    End If
                End If
            Next i
        
        'Favoriten in Array eintragen
            g_arr_varHorses(m_byteFavourite(1), 16) = 1
            g_arr_varHorses(m_byteFavourite(2), 16) = 2
            g_arr_varHorses(m_byteFavourite(3), 16) = 3
    
    'Wenn Taktik aktiviert
        If objOption.TACTICS > 0 Then 'Geschwindigkeit in den Rennphasen 1-3 pro Pferd festlegen
            For i = 1 To objRace.NUMBER_ENROLLED
                Randomize 'Zufallsgenerator zurücksetzen
                k = (Int((6 - 1 + 1) * Rnd + 1)) 'Zufallszahl zwischen 1 und 6
                For j = 1 To 3
                    g_arr_varHorses(i, 11 + j) = m_wksTec.Cells(1 + j, 3 + k).Value
                Next j
            Next i
        End If
        
        If objOption.TACTICS = 6 Then 'Geschwindigkeit in den Rennphasen 1-6 pro Pferd festlegen
            For i = 1 To objRace.NUMBER_ENROLLED
                Randomize 'Zufallsgenerator zurücksetzen
                k = (Int((6 - 1 + 1) * Rnd + 1)) 'Zufallszahl zwischen 1 und 6
                For j = 1 To 3
                    g_arr_varHorses(i, 18 + j) = m_wksTec.Cells(1 + j, 3 + k).Value
                Next j
            Next i
        End If
        
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

Public Sub ResetExcelOptions()
    Call SetExcelOptions(True, True, True, _
                            True, True, True, _
                            True, False, True)
End Sub

Public Sub ExcelOptionsTVmenu()
    Call SetExcelOptions(False, False, False, _
                             False, False, False, _
                             False, False, True)
End Sub

Public Sub ExcelOptionsTVfull()
    Call SetExcelOptions(False, False, False, _
                             False, False, False, _
                             False, True, True)
End Sub

Private Sub SetExcelOptions(blnGrid As Boolean, blnHead As Boolean, blnFormula As Boolean, _
                            blnStatus As Boolean, blnVScroll As Boolean, blnHScroll As Boolean, _
                            blnTabs As Boolean, blnFull As Boolean, blnMax As Boolean)

    With Application
        'Since some parameters depend on each other, the order of execution is important
'        If g_strPlayMode = "RS" Then
            .DisplayFullScreen = blnFull 'Excel ribbon
'        End If
            .ActiveWindow.DisplayGridlines = blnGrid  'Excel gridlines
            .ActiveWindow.DisplayHeadings = blnHead  'Excel row and column headings
            .DisplayFormulaBar = blnFormula  'Excel formula bar
            .DisplayStatusBar = blnStatus  'Excel status bar
            .ActiveWindow.DisplayVerticalScrollBar = blnVScroll  'Excel vertical scrollbar
            .ActiveWindow.DisplayHorizontalScrollBar = blnHScroll  'Excel horizontal scrollbar
            .ActiveWindow.DisplayWorkbookTabs = blnTabs  ' Excel workbook tabs
        If blnMax = True Then 'Window size
            .ActiveWindow.WindowState = xlMaximized
        Else
            .ActiveWindow.WindowState = xlNormal
        End If
    End With

End Sub

Public Function ZoomLevelPictures() As Variant() 'return an array with the values for the column width and row height
        Dim dblWindowHeight As Double
        
        dblWindowHeight = Application.ActiveWindow.Height 'window height
        
        Select Case dblWindowHeight
            Case Is > 1100 'large window
                ZoomLevelPictures = Array(3, 22) 'column width, row height
            Case Is > 799 'window height
                ZoomLevelPictures = Array(2, 15) 'column width, row height
            Case Else 'small window
                ZoomLevelPictures = Array(1.8, 14) 'column width, row height
        End Select
End Function

Public Function ZoomLevelRecommendation() As Byte 'Determination of the zoom level recommendation

    Dim dblWindowHeight As Double
    Dim intHorses As Byte

    dblWindowHeight = Application.ActiveWindow.Height 'window height
    
    'Get the number of horses in the selected race
    intHorses = ThisWorkbook.Worksheets(objRace.SELECTED).Cells( _
        basAuxiliary.GetRow(ThisWorkbook.Worksheets(objRace.SELECTED), "NUMBER OF STARTERS"), 2).Value

    Select Case dblWindowHeight
        Case Is > 790 'window height (e.g. 1680x1050 screen)
            If intHorses <= 23 Then 'up to 23 horses
                ZoomLevelRecommendation = 3
            ElseIf intHorses <= 31 Then '24-31 horses
                ZoomLevelRecommendation = 2
            Else 'more than 31 horses
                ZoomLevelRecommendation = 1
            End If
        Case Is > 550 'window height (e.g. 1368x768 screen)
            If intHorses <= 9 Then 'up to 9 horses
                ZoomLevelRecommendation = 3
            ElseIf intHorses <= 14 Then '10-14 horses
                ZoomLevelRecommendation = 2
            Else 'more than 14 horses
                ZoomLevelRecommendation = 1
            End If
        Case Else 'small window
            ZoomLevelRecommendation = 1
    End Select
End Function

'Retrieve zoom level text
Public Function ZoomLevelText(byteZL As Byte) As String
    Select Case byteZL
        Case 1
            ZoomLevelText = GetText(g_arr_Text, "ZOOM003")
        Case 2
            ZoomLevelText = GetText(g_arr_Text, "ZOOM004")
        Case 3
            ZoomLevelText = GetText(g_arr_Text, "ZOOM005")
    End Select
End Function

'Retrieve horse size preview
Public Function HorseSizePreview(intHSP) As Variant() 'return an array with the values
    Select Case intHSP
        Case 1
            HorseSizePreview = Array(12, 6) '(width, height) of the horse // small
        Case 2
            HorseSizePreview = Array(22, 9) '(width, height) of the horse // medium
        Case 3
            HorseSizePreview = Array(30, 12) '(width, height) of the horse // large
    End Select
End Function

Public Sub AnimalGrammar()
    Dim combo As String 'Race participants as chosen in the dropdown
                        '(HORSE/PIG/DONKEY/DOG/UNICORN)
    Dim col As Integer

    With ThisWorkbook
        combo = .Worksheets(objRace.SELECTED).Cells(basAuxiliary.GetRow(.Worksheets(objRace.SELECTED), "PARTICIPANTS"), 2).Value
    End With
    
    'Get the column with the language
    col = basAuxiliary.GetColumn(g_wksTEXT, objOption.language)

    For i = 1 To g_wksTEXT.Cells(rows.count, 1).End(xlUp).row 'ToDo... geht das in AI oder hart auf 1000?
        If g_wksTEXT.Cells(i, 1).Value = combo Then
            g_arr_Grammar(g_wksTEXT.Cells(i, 2).Value) = g_wksTEXT.Cells(i, col).Value
        End If
    Next i
    
End Sub

Private Sub DrawRaceTrack()
    
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    'Deactivate screen updating
        Application.ScreenUpdating = False
        
    'Freeze columns A-K if one of those checkboxes is activated, otherwise unfreeze
        If objOption.NAMES_LEFT Or objOption.COLOURS_LEFT Or objOption.HIGHLIGHT_FAV _
            Or (objOption.FOCUSED_RUN And objOption.HIGHLIGHT_FOC) Or (objOption.RACE_INFO And objOption.RACE_INFO_WKS) Then
                Call basAuxiliary.Freeze(12, 0, True) 'freeze
        Else
                Call basAuxiliary.Freeze(0, 0, False) 'unfreeze
        End If

    'Formatting: row height of the different sections
        g_wksRace.Range(rows(1 + m_intTopRows), rows(5 + m_intTopRows)).EntireRow.RowHeight = 15 'above the race track
        g_wksRace.Range(rows(objRace.NUMBER_ENROLLED * 2 + 6 + 1 + m_intTopRows), _
            rows(objRace.NUMBER_ENROLLED * 2 + 52 + m_intTopRows)).EntireRow.RowHeight = 15 'below the race track
        g_wksRace.rows(objRace.NUMBER_ENROLLED * 2 + 20 + m_intTopRows).RowHeight = 20 'headline of the ranking list
    'Display race data on the top
        With g_wksRace.Cells(2 + m_intTopRows, 14)
            .Font.name = "Arial Black"
            .Value = objRace.RACE_NAME & " " & objRace.RACE_YEAR & " - " & objRace.TRACK_NAME & ", " & objRace.TRACK_LOCATION _
                    & " (" & objRace.COUNTRY & ")"
        End With
        With g_wksRace.Cells(3 + m_intTopRows, 14)
            .Font.name = "Arial"
            .Font.Bold = True 'Fettschrift
            .Value = objRace.RACE_TYPE_TEXT & " " & GetText(g_arr_Text, "RACE007") & " " & _
                objRace.METERS & GetText(g_arr_Text, "RACE008") & " - " & objRace.TRACK_SURFACE_TEXT
        End With
    'Formatting: columns on the left side of the starting grid dependent on the zoom level
        With g_wksRace
            .Columns(1).ColumnWidth = 2 * objOption.ZOOM_LEVEL 'left margin
            .Range(Columns(2), Columns(9)).ColumnWidth = m_dblTrackCellWidth 'Horse colours
            .Columns(10).ColumnWidth = objOption.ZOOM_LEVEL 'empty column
            .Columns(11).ColumnWidth = 2 * objOption.ZOOM_LEVEL 'horse numbers
            .Columns(12).ColumnWidth = 20 + (objOption.ZOOM_LEVEL * 6) 'horse names
            .Range(Columns(13), Columns(m_intLeftColumns + 12)).ColumnWidth = m_dblTrackCellWidth 'cell length of the race track
            .Columns(m_intLeftColumns + 4).ColumnWidth = 3 + objOption.ZOOM_LEVEL 'starting box numbers
    'Formatting: race track
            .Range(Columns(m_intLeftColumns + 13), Columns(objRace.METERS + m_intLeftColumns + 13 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))).ColumnWidth = m_dblTrackCellWidth  'cell length of the race track
        End With
        
        'Cell height of the race track
        For i = (6 + m_intTopRows) To (objRace.NUMBER_ENROLLED * 2 + 6 + m_intTopRows)
            g_wksRace.rows(i).RowHeight = m_intTrackCellHeight / (2 - (i - (6 + m_intTopRows)) Mod 2) 'alternate...
        Next i
        
        'Track colour
        g_wksRace.Range(Cells(4 + m_intTopRows, 1), Cells(objRace.NUMBER_ENROLLED * 2 + 19 + m_intTopRows, objRace.METERS + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))).Interior.color = objRace.TRACK_COLOUR
        
        'Track font
        With g_wksRace.Range(Cells(4 + m_intTopRows, 1), Cells(objRace.NUMBER_ENROLLED * 2 + 8 + m_intTopRows, objRace.METERS + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)))
            .Font.name = "Arial"
            .Font.Size = m_intFontSize
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    
    If objRace.STARTING_GATE = "Y" Then
        'Draw the start boxes
            For i = 6 To (objRace.NUMBER_ENROLLED * 2 + 6) Step 2 'Für jeden Startplatz eine Box
                g_wksRace.Range(Cells(i + m_intTopRows, m_intLeftColumns + 8), Cells(i + m_intTopRows, m_intLeftColumns + 13)).Interior.ColorIndex = 1 'schwarz
            Next i
            g_wksRace.Range(Cells(6 + m_intTopRows, m_intLeftColumns + 13), Cells(objRace.NUMBER_ENROLLED * 2 + 6 + m_intTopRows, m_intLeftColumns + 13)).Interior.ColorIndex = 1 'schwarz
        'Label the start boxes
            With g_wksRace.Range(Cells(7 + m_intTopRows, m_intLeftColumns + 4), Cells(objRace.NUMBER_ENROLLED * 2 + 5 + m_intTopRows, m_intLeftColumns + 4))
                .Font.ColorIndex = 1 'colour: black
                .Font.Size = m_intFontSize
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlCenter
            End With
            For i = 1 To (objRace.NUMBER_ENROLLED) 'one start box for each horse
                g_wksRace.Cells(5 + 2 * i + m_intTopRows, m_intLeftColumns + 4).Value = GetText(g_arr_Text, "RACE010") & " " & i
            Next i
    End If
        
    'Display meters above and below the race track
        For i = objOption.METRES_DISPLAY To (objRace.METERS - 20) Step objOption.METRES_DISPLAY
            'For debugging purposes
            #If Debugging Then
                g_wksRace.Range(Cells(5, i + m_intLeftColumns + 11), Cells(45, i + m_intLeftColumns + 11)).Interior.ColorIndex = 1
            #End If
            With g_wksRace.Cells(4 + m_intTopRows, i + m_intLeftColumns + 11)
                .Font.name = "Arial"
                .Font.Bold = True 'Fettschrift
                .Value = i & GetText(g_arr_Text, "RACE008") '"m"
            End With
            With g_wksRace.Cells(objRace.NUMBER_ENROLLED * 2 + 8 + m_intTopRows, i + m_intLeftColumns + 11)
                .Font.name = "Arial"
                .Font.Bold = True 'Fettschrift
                .Value = i & GetText(g_arr_Text, "RACE008") '"m"
            End With
        Next i
    'Formatting: horse names on the left
        With g_wksRace.Range(Cells(6 + m_intTopRows, 11), Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, 12))
            .Font.color = objRace.TRACK_COLOUR  'track colour, so that the names are not visible
            .IndentLevel = 1 'text indented
            .Font.Size = m_intFontSize
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        
    'In case of a mudflats race
        If objRace.TRACK_SURFACE = "M" Then Call DrawMudflats
        
    'Formatting: finish area
        g_wksRace.Columns(objRace.METERS + m_intLeftColumns + 12).ColumnWidth = m_dblTrackCellWidth 'width of the finish line
        g_wksRace.Range(Cells(5 + m_intTopRows, objRace.METERS + m_intLeftColumns + 12), _
            Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METERS + m_intLeftColumns + 12)).Interior.ColorIndex = 56 'colour of the finish line: dark grey
        With g_wksRace.Cells(4 + m_intTopRows, objRace.METERS + m_intLeftColumns + 11)
            .Font.name = "Arial"
            .Font.Bold = True
            .Value = objRace.METERS & GetText(g_arr_Text, "RACE008") 'distance in meters
        End With
        With g_wksRace.Cells(objRace.NUMBER_ENROLLED * 2 + 8 + m_intTopRows, objRace.METERS + m_intLeftColumns + 11)
            .Font.name = "Arial"
            .Font.Bold = True
            .Value = objRace.METERS & GetText(g_arr_Text, "RACE008") 'distance in meters
        End With
    'Area behind the finish line
        g_wksRace.Columns(objRace.METERS + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).ColumnWidth = 18 + (objOption.ZOOM_LEVEL * 6)
        g_wksRace.Columns(objRace.METERS + m_intLeftColumns + 15 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).ColumnWidth = objOption.ZOOM_LEVEL * 3
    'Formatting: horse names in the finish
        With g_wksRace.Range(Cells(5 + m_intTopRows, objRace.METERS + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), Cells(objRace.NUMBER_ENROLLED * 2 + 7 + (2 * objOption.SPEED_FACTOR) + m_intTopRows, objRace.METERS + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)))
            .Font.ColorIndex = 1 'colour: black
            .IndentLevel = 1 'text indented
            .Font.Size = m_intFontSize
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    
    'Advertising below the race track
        g_wksRace.Range(rows(objRace.NUMBER_ENROLLED * 2 + 9 + m_intTopRows), _
            rows(objRace.NUMBER_ENROLLED * 2 + 19 + m_intTopRows)).EntireRow.RowHeight = m_intAdvertisingHeight 'row height according to the zoom level

        If objRace.ADVERTISING = "Y" Then
            Dim advPos As Integer 'column position for next ad
            
            'Get data from the race sheet
            Call GetAdvertisementData
    
            advPos = m_intLeftColumns + 12
            For i = 1 To UBound(m_arr_varAdv) 'draw the advertising
                z = 3 'set the pointer to the first colour code
                For j = advPos To advPos + g_wksPIC.Cells(2, m_arr_varAdv(i)) - 1
                    If j >= objRace.METERS + m_intLeftColumns + 12 Then Exit For
                    For k = objRace.NUMBER_ENROLLED * 2 + 9 + m_intTopRows To objRace.NUMBER_ENROLLED * 2 + 19 + m_intTopRows
                        g_wksRace.Cells(k, j).Interior.color = g_wksPIC.Cells(z, m_arr_varAdv(i)).Value
                        z = z + 1
                    Next k
                Next j
                advPos = advPos + g_wksPIC.Cells(2, m_arr_varAdv(i))
            Next i
        End If
    
    'Formatting: finish photo (headline)
        g_wksRace.Cells(2 + m_intTopRows, objRace.METERS + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Font.name = "Arial Black"  '"FOTOFINISH!"
        With g_wksRace.Cells(4 + m_intTopRows, objRace.METERS + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)) '"Zielfoto..."
            .Font.Size = 14
            .Font.Bold = True
        End With

    'Formatting: finish photo and ranking list
        g_wksRace.Range(Columns(objRace.METERS + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), Columns(objRace.METERS + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))).ColumnWidth = m_dblRankingsWidth / 10   'column width according to the zoom level
        g_wksRace.Range(Columns(objRace.METERS + m_intLeftColumns + 176 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR)), Columns(objRace.METERS + m_intLeftColumns + 176 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))).ColumnWidth = m_dblRankingsWidth   'column width according to the zoom level
        'Formatting: photo of the winner
        g_wksRace.Range(Columns(objRace.METERS + m_intLeftColumns + 177 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR)), Columns(objRace.METERS + m_intLeftColumns + 197 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))).ColumnWidth = 2
        With g_wksRace.Range(Cells(objRace.NUMBER_ENROLLED * 2 + 20 + m_intTopRows, objRace.METERS + m_intLeftColumns + 177 + (2 * 10 * objOption.SPEED_FACTOR)), _
                        Cells(objRace.NUMBER_ENROLLED * 2 + 21 + m_intTopRows, objRace.METERS + m_intLeftColumns + 179 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR)))
            .Font.Size = 14
            .Font.Bold = True
        End With
    
    'Place the cursor far away
        g_wksRace.Cells(100 + m_intTopRows, 1).Select
    'Activate screen updating
        Application.ScreenUpdating = True
        
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Track for a mudflats race
Private Sub DrawMudflats()
    'For debugging purposes: Lugworm population
    #If Debugging Then
        Dim lngLug1 As Long, lngLug2 As Long
    #End If
        
        Dim rndPuddleFrequency As Integer
        Dim rndPuddleLength As Integer
        Dim rndPuddleWidth As Integer
    
        Dim rndLugwormFrequency As Integer
        Dim rndLugwormShape As Integer
        
        Dim c As Integer 'column
        
    'Draw Lugworms
    If objOption.LUGWORMS > 0 Then
        objOption.LUGWORM_COL = 2770764 'Colour of the lugworms in the Wadden Sea
        c = basAuxiliary.GetColumn(g_wksPIC, "LUGWORMS")
        ReDim m_lngLugworms(1 To (g_wksPIC.Cells(rows.count, c).End(xlUp).row - 1)) 'ToDo... geht das in AI???
        
        For i = 1 To UBound(m_lngLugworms)
            m_lngLugworms(i) = g_wksPIC.Cells(i + 1, c) 'Get the lugworm characters
        Next i

        For i = (5 + m_intTopRows) To (objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows) 'rows
            For j = m_intLeftColumns + 15 To objRace.METERS + m_intLeftColumns + 7 'columns
                rndLugwormFrequency = Int(((100 / objOption.LUGWORMS) - 1 + 1) * Rnd + 1) 'lugworm or no lugworm
                If rndLugwormFrequency = 1 Then
                    rndLugwormShape = Int((UBound(m_lngLugworms) - 1 + 1) * Rnd + 1) 'lugworm shape
                    With g_wksRace.Cells(i, j)
                        .Font.color = objOption.LUGWORM_COL
                        .Value = ChrW(m_lngLugworms(rndLugwormShape))
                    End With
                    #If Debugging Then
                        lngLug1 = lngLug1 + 1
                    #End If
                End If
                #If Debugging Then
                    lngLug2 = lngLug2 + 1
                #End If
            Next j
        Next i
        #If Debugging Then
            Debug.Print lngLug1 & " lugworms (" & lngLug1 / lngLug2 * 100 & "%)"
        #End If
    End If

    'Draw Puddles
    If objOption.TIDE > 0 Then
        objOption.PUDDLE_COL = 10791854 'Colour of the puddles in the Wadden Sea
        
        For i = (5 + m_intTopRows) To (objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows) 'rows
            For j = m_intLeftColumns + 15 To objRace.METERS + m_intLeftColumns + 7 'columns
                rndPuddleFrequency = Int(((100 / objOption.TIDE) - 1 + 1) * Rnd + 1) 'puddle or no puddle
                
                If rndPuddleFrequency = 1 Then
                    rndPuddleLength = Int((10 - 1 + 1) * Rnd + 1) 'puddle length
                    rndPuddleWidth = Int((2 - 1 + 1) * Rnd + 1) 'puddle width
                
                    With g_wksRace.Range(Cells(i, j), Cells(i + rndPuddleWidth - 1, j + rndPuddleLength - 1))
                        .Interior.color = objOption.PUDDLE_COL
                        .Font.color = objOption.PUDDLE_COL
                        .Value = "|"
                    End With
                    
                    #If Debugging Then
                        g_wksRace.Range(Cells(i, j), Cells(i + rndPuddleWidth - 1, j + rndPuddleLength - 1)) _
                            .Font.color = vbBlack
                    #End If

                End If

            Next j
        Next i
    End If
End Sub

'Draw horse names during the race
Private Sub DrawHorseNames()
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    'Pferdenamen am linken Rand platzieren wenn Checkbox angehakt ist
        If objOption.NAMES_LEFT Then
            For i = 1 To objRace.NUMBER_ENROLLED
                If g_arr_varHorses(i, 0) = "START" Then
                    g_wksRace.Cells(g_arr_varHorses(i, 3), 11).Value = "#" & g_arr_varHorses(i, 11)  'Startnummer
                    g_wksRace.Cells(g_arr_varHorses(i, 3), 12).Value = g_arr_varHorses(i, 1)  'Name des Pferds
                End If
            Next i
        'Optimale Spaltenbreite
            g_wksRace.Range(Columns(10), Columns(11)).EntireColumn.AutoFit
        End If
    'Pferdenamen im Ziel anzeigen wenn Checkbox angehakt ist
        If objOption.NAMES_FINISH Then
            For j = 1 To objRace.NUMBER_ENROLLED
                If g_arr_varHorses(j, 0) = "START" Then
                    g_wksRace.Cells(g_arr_varHorses(j, 3), objRace.METERS + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = _
                        g_arr_varHorses(j, 1) & " (#" & g_arr_varHorses(j, 11) & ")"
                End If
            Next j
        End If
        
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Race information pop-up
Private Sub RaceInfoPopup()
    With frmRaceInfo
        'Place the pop-up in the upper left corner
            .StartUpPosition = 0
            .top = ActiveWindow.top + 20
            .left = ActiveWindow.left + 20
        .BackColor = objOption.RACE_INFO_COL_B
        .Caption = GetText(g_arr_Text, "USERFORM006")

        'Labels mit Name und Distanz des Rennens (ToDoo...auch auslagern in clsModul!)
        With .lbl_RI1
            .BackColor = objOption.RACE_INFO_COL_B
            .ForeColor = objOption.RACE_INFO_COL_F
            .Caption = objRace.RACE_NAME & " " & objRace.RACE_YEAR
            .Font.Size = 12
            .Font.Bold = True
            .AutoSize = True
        End With
        With .lbl_RI2
            .BackColor = objOption.RACE_INFO_COL_B
            .ForeColor = objOption.RACE_INFO_COL_F
            .Font.Size = 12
            .Caption = GetText(g_arr_Text, "RACE024") & ": " & objRace.METERS & GetText(g_arr_Text, "RACE008")
            .AutoSize = True
        End With
        
        'ToDooo... Auslagern in clsModul (149.00)
        If objOption.RACE_INFO_PROGRESS Then
            Set g_objLabel = .Controls.Add("Forms.Label.1", , True)
            With g_objLabel 'progress bar (background)
                .name = "lbl_RI3a_dyn"
                .Font.name = "Tahoma"
                .Font.Size = 8
                .left = 6
                .top = frmRaceInfo.Height - 25
                .Width = 200
                .Height = 12
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = objOption.RACE_INFO_COL_F
                .ForeColor = objOption.RACE_INFO_COL_F
                .BackColor = objOption.RACE_INFO_COL_B
                .TextAlign = fmTextAlignRight
                .Caption = objRace.METERS
            End With
            Set g_objLabel = .Controls.Add("Forms.Label.1", , True)
            With g_objLabel 'progress bar (bar)
                .name = "lbl_RI3b_dyn"
                .Font.name = "Tahoma"
                .Font.Size = 8
                .left = 6
                .top = frmRaceInfo.Height - 25
                .Width = 0
                .Height = 12
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = objOption.RACE_INFO_COL_F
                .ForeColor = objOption.RACE_INFO_COL_B
                .BackColor = objOption.RACE_INFO_COL_F
                .TextAlign = fmTextAlignLeft
            End With
            'Adjust UserForm height
            frmRaceInfo.Height = frmRaceInfo.Height + g_objLabel.Height + 6
        End If

        If objOption.RACE_INFO_LEADER Then
            Set g_objLabel = .Controls.Add("Forms.Label.1", , True)
            With g_objLabel 'name of the leader
                .name = "lbl_RI4a_dyn"
'                        .Font.name = "Tahoma"
                .Font.Size = 12
                .left = 6
                .top = frmRaceInfo.Height - 25
                .Width = 200
                .Height = 18
                .ForeColor = objOption.RACE_INFO_COL_F
                .TextAlign = fmTextAlignLeft
                .Caption = ""
            End With
            'Adjust UserForm height
            frmRaceInfo.Height = frmRaceInfo.Height + g_objLabel.Height
            Set g_objLabel = .Controls.Add("Forms.Label.1", , True)
            With g_objLabel 'name of the leader
                .name = "lbl_RI4b_dyn"
'                        .Font.name = "Tahoma"
                .Font.Size = 12
                .left = 6
                .top = frmRaceInfo.Height - 25
                .Width = 200
                .Height = 18
                .ForeColor = objOption.RACE_INFO_COL_F
                .TextAlign = fmTextAlignCenter
                .Caption = ""
            End With
            'Adjust UserForm height
            frmRaceInfo.Height = frmRaceInfo.Height + g_objLabel.Height + 6
        End If

        .show (vbModeless) 'modeless
    End With
End Sub

'Meldung zum Rennstart
Private Sub RaceWelcome()
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    'A MessageBox cannot handle unicode, so for example Cyrillic characters are displayed as question marks
'    MsgBox GetText(g_arr_Text, "RACE001") & " " & GetText(g_arr_Text, "RACE006") & " " & objRace.TRACK_LOCATION & " (" & objRace.COUNTRY & "). " & vbNewLine & vbNewLine & _
'            GetText(g_arr_Text, "RACE003") & " " & m_strRaceName & " " & GetText(g_arr_Text, "RACE007") & " " & objRace.METERS & " " & GetText(g_arr_Text, "RACE009") & "." & vbNewLine & vbNewLine & _
'            GetText(g_arr_Text, "RACE004") & " " & m_intHorsesStarting & " " & g_arr_Grammar(4) & ".", , g_c_tool
    
    'Set the button mode
    g_strMsgButtons = "OK"
    'Assign the text for the pop-up
    g_strMsgCaption = g_c_tool
    g_strMsgText = GetText(g_arr_Text, "RACE001") & " " & GetText(g_arr_Text, "RACE006") & " " & objRace.TRACK_LOCATION & " (" & objRace.COUNTRY & "). " & vbNewLine & vbNewLine & _
            GetText(g_arr_Text, "RACE003") & ": " & objRace.RACE_NAME & " " & GetText(g_arr_Text, "RACE007") & " " & objRace.METERS & " " & GetText(g_arr_Text, "RACE009") & "." & vbNewLine & vbNewLine & _
            GetText(g_arr_Text, "RACE004") & " " & m_intHorsesStarting & " " & g_arr_Grammar(4) & "."
    
    If objOption.SPEECH Then Call SpeechOut(g_strMsgText)
    
    'Display the pop-up
    frmMsg_MultiPurpose.show (vbModal) 'modal
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Pferde in Boxen stellen
Private Sub StartingGrid()
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    If Not DEVMODE Then Application.wait (Now + TimeValue("0:00:02")) 'Verzögerung
    
    For i = 1 To objRace.NUMBER_ENROLLED
        If g_arr_varHorses(i, 0) = "START" Then
            Call PaintHorse(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7, g_arr_varHorses(i, 2))
        End If
    Next i
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Presentation of the horses with numbers and names
Private Sub RacePresentation()
    'In case of a runtime error
    On Error GoTo ERRORHANDLING 'REIN
    
    If Not DEVMODE Then Application.wait (Now + TimeValue("0:00:02"))     'delay
    
    Application.DisplayCommentIndicator = xlCommentAndIndicator 'turn on cell comments
    
    'Display a comment field for each horse
    For i = 1 To objRace.NUMBER_ENROLLED
        If g_arr_varHorses(i, 0) = "START" Then
            With g_wksRace.Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4))
                If i = m_byteFavourite(1) Then 'enhance comment field
                    If objOption.FOCUSED_RUN And g_arr_varHorses(i, 11) = objOption.FOCUSED_NR Then
                        .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) _
                            & " (" & GetText(g_arr_Text, "RACE011") & ") >> " & GetText(g_arr_Text, "RACE012") 'horse number, name and (favourite) >> in focus
                    Else
                        .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) _
                            & " (" & GetText(g_arr_Text, "RACE011") & ")" 'horse number, name and (favourite)
                    End If
                ElseIf objOption.FOCUSED_RUN And g_arr_varHorses(i, 11) = objOption.FOCUSED_NR Then
                    .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) _
                        & " >> " & GetText(g_arr_Text, "RACE012") 'horse number, name and >> in focus
                Else
                    .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) 'horse number and name
                End If
                .Comment.Shape.TextFrame.Characters.Font.Size = m_intFontSize 'font size according to the zoom level
                .Comment.Shape.TextFrame.AutoSize = True 'optimise the size of the comment field
            End With
            If i = m_byteFavourite(1) Then 'highlight the favourite horse
                g_wksRace.Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)) _
                    .Comment.Shape.Fill.ForeColor.RGB = RGB(192, 0, 0) 'red background
            End If
            'Draw a blue dashed frame around the horse on focus if the checkbox is activated
                If objOption.FOCUSED_RUN Then
                    If g_arr_varHorses(i, 11) = objOption.FOCUSED_NR Then
                        g_wksRace.Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
                            Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)) _
                            .BorderAround ColorIndex:=41, LineStyle:=xlDash, Weight:=xlThick
                    End If
                End If
        End If
    Next i

    'Announce the favourite horses
        'Set the button mode
        g_strMsgButtons = "OK"
        'Assign the text for the pop-up
        g_strMsgCaption = objRace.RACE_NAME & " " & objRace.RACE_YEAR
        g_strMsgText = GetText(g_arr_Text, "RACE013") & " " & g_arr_varHorses(m_byteFavourite(1), 1) & _
                " (#" & g_arr_varHorses(m_byteFavourite(1), 11) & ")." & vbNewLine & vbNewLine & _
                GetText(g_arr_Text, "RACE015") & " " & g_arr_varHorses(m_byteFavourite(2), 1) & " (#" _
                & g_arr_varHorses(m_byteFavourite(2), 11) & ") " & vbNewLine & _
                GetText(g_arr_Text, "RACE017") & " " & g_arr_varHorses(m_byteFavourite(3), 1) & " (#" & _
                g_arr_varHorses(m_byteFavourite(3), 11) & ") " & GetText(g_arr_Text, "RACE018") & "."
        
        If objOption.SPEECH Then Call SpeechOut(g_strMsgText)
        
        'Display the pop-up
        frmMsg_MultiPurpose.show (vbModal) 'modal

    'Announce the focused horse
        If objOption.FOCUSED_RUN Then
            For i = 1 To UBound(g_arr_varHorses)
                If g_arr_varHorses(i, 11) = objOption.FOCUSED_NR Then
                    'Set the button mode
                    g_strMsgButtons = "OK"
                    'Assign the text for the pop-up
                    g_strMsgCaption = GetText(g_arr_Text, "RACEOPT026") 'Focused Run
                    g_strMsgText = GetText(g_arr_Text, "RACE021") & " " & g_arr_varHorses(i, 1) & " (#" & g_arr_varHorses(i, 11) & ")."

                    If objOption.SPEECH Then Call SpeechOut(g_strMsgText)
        
                    'Display the pop-up
                    frmMsg_MultiPurpose.show (vbModal) 'modal
                    Exit For
                End If
            Next i
        End If
                
    'Turn off cell comments (hide the horse names)
        Application.DisplayCommentIndicator = xlNoIndicator
        
    'Farben der Pferde am linken Rand platzieren wenn Checkbox angehakt ist
        If objOption.COLOURS_LEFT Then
            For i = 1 To objRace.NUMBER_ENROLLED
                If g_arr_varHorses(i, 0) = "START" Then
                    Call PaintHorse(g_arr_varHorses(i, 3), 2, g_arr_varHorses(i, 2))
                End If
            Next i
        End If
        
    'Pferdenamen und Startnummern am linken Rand zeigen wenn Checkbox angehakt ist
        g_wksRace.Range(Columns(11), Columns(12)).Font.ColorIndex = 1 'schwarz
    
    'Mark the favourite on the left if the checkbox is ticked
        If objOption.HIGHLIGHT_FAV Then
            g_wksRace.Range(Cells(g_arr_varHorses(m_byteFavourite(1), 3), 11), Cells(g_arr_varHorses(m_byteFavourite(1), 3), 12)) _
                .Interior.color = 192 'background color: red
            'Show horse data on the left
            Call PaintHorse(g_arr_varHorses(m_byteFavourite(1), 3), 2, g_arr_varHorses(m_byteFavourite(1), 2))
            g_wksRace.Cells(g_arr_varHorses(m_byteFavourite(1), 3), 11).Value = "#" & g_arr_varHorses(m_byteFavourite(1), 11) 'horse number
            g_wksRace.Cells(g_arr_varHorses(m_byteFavourite(1), 3), 12).Value = g_arr_varHorses(m_byteFavourite(1), 1) _
                    & " (" & GetText(g_arr_Text, "RACE011") & ")" 'horse name
        End If
        
    'Adapt the frame around the focused horse
        If objOption.FOCUSED_RUN Then
            For i = 1 To UBound(g_arr_varHorses)
                If g_arr_varHorses(i, 11) = objOption.FOCUSED_NR Then
                    'Delete the frame around the horse
                    g_wksRace.Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
                        Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)) _
                        .Borders.LineStyle = xlLineStyleNone
                    'Highlight the focused horse on the left if the checkbox is ticked
                    If objOption.HIGHLIGHT_FOC Then
                        'Draw a new frame
                        g_wksRace.Range(Cells(g_arr_varHorses(i, 3), 11), Cells(g_arr_varHorses(i, 3), 12)) _
                            .BorderAround ColorIndex:=41, LineStyle:=xlDash, Weight:=xlThick
                        'Show horse data on the left
                        Call PaintHorse(g_arr_varHorses(i, 3), 2, g_arr_varHorses(i, 2)) 'Horse colour
                        g_wksRace.Cells(g_arr_varHorses(i, 3), 11).Value = "#" & g_arr_varHorses(i, 11) 'Horse number
                        g_wksRace.Cells(g_arr_varHorses(i, 3), 12).Value = g_wksRace.Cells(g_arr_varHorses(i, 3), 12).Value _
                                & " >> " & GetText(g_arr_Text, "RACE012") ' >> in focus
                    End If
                    Exit For
                End If
            Next i
        End If
    
    'Delay before the start
     If Not DEVMODE Then Application.wait (Now + TimeValue("0:00:04"))

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Start des Galopprennens
Private Sub RunRace()

    Dim strDistance As String '"Race distance"
    Dim strMeter As String '"m"
    Dim strLeader '"The current leader is"
    Dim intPipesProgressBar As Integer 'number of pipes of the worksheet progress bar dependent on the zoom level
    Dim intRefuse As Integer 'für Zufallsgenerator
    Dim intSquirtPattern As Integer 'für Zufallsgenerator
    Dim intSquirtLength As Integer 'für Zufallsgenerator
    Dim dblSquirtColour As Double 'für Zufallsgenerator
    
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    m_intHorsesRunning = m_intHorsesStarting 'initial number of starters
    
    'One out of 100 refuses to run
    If objOption.REFUSE_RUN Then
        For i = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(i, 0) = "START" Then
                Randomize 'Zufallsgenerator zurücksetzen
                intRefuse = Int((99 - 0 + 1) * Rnd + 0)
                If intRefuse = 0 Then
                    g_arr_varHorses(i, 0) = "REFUSED"
                    m_intHorsesRunning = m_intHorsesRunning - 1
                End If
            End If
        Next i
    End If
    
    'Preparing race information on the worksheet
        'Calculation of the number of pipes of the progress bar
        If objOption.RACE_INFO And objOption.RACE_INFO_WKS Then
        'TODO: Check race progress bar
'           intPipesProgressBar = 32 + (objOption.ZOOM_LEVEL * 8) 'Duncan
           intPipesProgressBar = 4 + ((3 + objOption.ZOOM_LEVEL) * 14) 'Marco
        End If
        'Formatting
        If objOption.RACE_INFO And objOption.RACE_INFO_WKS Then
            Call basAuxiliary.RaceInfoWorksheet(objOption.RACE_INFO_COL_B, objOption.RACE_INFO_COL_F, m_intTopRows, True)
        End If
        
    If objRace.STARTING_GATE = "Y" Then
        'Boxennummern entfernen
            g_wksRace.Range(Cells(7 + m_intTopRows, m_intLeftColumns + 4), Cells(5 + 2 * objRace.NUMBER_ENROLLED + m_intTopRows, m_intLeftColumns + 4)).Value = ""
        'Verzögerung
            If Not DEVMODE Then Application.wait (Now + TimeValue("0:00:04"))
        'Boxen aufmachen
            g_wksRace.Range(Cells(6 + m_intTopRows, m_intLeftColumns + 13), Cells(objRace.NUMBER_ENROLLED * 2 + 6 + m_intTopRows, m_intLeftColumns + 13)).Interior.color = objRace.TRACK_COLOUR
    End If
        
    'Speech
    If objOption.SPEECH Then Call SpeechOut(GetText(g_arr_Text, "RACE034"))
            
    'Race information data
        'Labels
        strDistance = GetText(g_arr_Text, "RACE024")
        strMeter = GetText(g_arr_Text, "RACE008")
        strLeader = GetText(g_arr_Text, "RACEINFO001")
        'Name and position of the leader
        m_intPosLeader = g_arr_varHorses(1, 4) - (m_intLeftColumns + 12) 'zero metres
        m_strNameLeader = ""
    'Fotofinish zurücksetzen
        m_blnPhotofinish = False
    'Noch kein Pferd ist im Ziel
        m_blnWin = False
        m_blnDeadHeat = False
        m_blnExactFinish = False
    'Platzierung für das nächste Pferd, das im Ziel ankommt
        m_intPlace = 1
    'Berechnungsrunde für die Platzierung zurücksetzen
        m_intFinishLoop = 0
    
    'Rennen läuft
        'For debugging purposes: Race start time
        #If Debugging Then
            Dim timeStart As Date
            timeStart = Now
            Debug.Print m_strRaceName & " (" & objRace.METERS & ")"
            Debug.Print "Race start : " & Format(timeStart, "HH:MM:SS")
        #End If
        
        Do Until m_intPlace > m_intHorsesRunning 'solange noch nicht alle im Ziel sind
            'Zählvariable für den Zieleinlauf pro Schleifendurchlauf zurücksetzen
                m_intHorsesFinishing = 0
            'Neue Positionen berechnen
                For i = 1 To UBound(g_arr_varHorses)
                    'Geschwindigkeitsfaktor pro Durchlauf
                    g_arr_varHorses(i, 7) = SpeedLoop()
                    'Schrittweite pro Durchlauf (ungerundet)
                        If objOption.TACTICS = 0 Then
                        'Wenn Geschwindigkeit pro Rennphase konstant sein soll
                            g_arr_varHorses(i, 8) = _
                                (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + g_arr_varHorses(i, 7)) / 3
                        ElseIf objOption.TACTICS = 3 Then
                        'Wenn jedes Pferd in einem Renndrittel unterschiedlich schnell sein soll
                            'Berechnen in welchem Streckenabschnitt das Pferd ist
                            Select Case True
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) < objRace.METERS * 1 / 3 'Pferd ist im 1. Renndrittel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 12)) / 4
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) < objRace.METERS * 2 / 3 'Pferd ist im 2. Renndrittel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 13)) / 4
                                Case Else 'Pferd ist im 3. Renndrittel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 14)) / 4
                                End Select
                        ElseIf objOption.TACTICS = 6 Then
                            'Berechnen in welchem Streckenabschnitt das Pferd ist
                            Select Case True
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) < objRace.METERS * 1 / 6 'Pferd ist im 1. Rennsechstel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 12)) / 4
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) < objRace.METERS * 2 / 6 'Pferd ist im 2. Rennsechstel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 13)) / 4
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) < objRace.METERS * 3 / 6 'Pferd ist im 3. Rennsechstel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 14)) / 4
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) < objRace.METERS * 4 / 6 'Pferd ist im 4. Rennsechstel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 19)) / 4
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) < objRace.METERS * 5 / 6 'Pferd ist im 5. Rennsechstel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 20)) / 4
                                Case Else 'Pferd ist im 6. Rennsechstel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 21)) / 4
                                End Select
                        End If
                        
                    'Wasserspritzer entfernen
                            If objRace.SQUIRT = True Then
                                    Range(Cells(g_arr_varHorses(i, 3) - 1, g_arr_varHorses(i, 4) - 6), _
                                        Cells(g_arr_varHorses(i, 3) + 1, g_arr_varHorses(i, 4) - 14)).Interior.Pattern = xlSolid
                            End If
                            
                    'Windschatten entfernen wenn Pferd drin
                        If objOption.SLIPSTREAM And objOption.SLIPSTREAM_SHOW And g_arr_varHorses(i, 22) > 0 Then
                            If g_arr_varHorses(i, 4) <= objRace.METERS + m_intLeftColumns + 9 Then
                                Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9), _
                                    Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 12)).Interior.Pattern = xlSolid
                            End If
                        End If
                        
                    'Windschatten zurücksetzen
                        g_arr_varHorses(i, 22) = 0
                    
                    'Windschatten berechnen
                        If objOption.SLIPSTREAM Then 'if slipstreaming is activated
                            For k = 1 To UBound(g_arr_varHorses) 'loop through the horses
                            
                                If g_arr_varHorses(i, 15) - 1 = g_arr_varHorses(k, 15) _
                                    Or g_arr_varHorses(i, 15) + 1 = g_arr_varHorses(k, 15) Then 'one box above or below
                                        If g_arr_varHorses(i, 4) > g_arr_varHorses(k, 4) - 8 _
                                            And g_arr_varHorses(i, 4) < g_arr_varHorses(k, 4) - 4 Then
                                                If objOption.SLIPSTREAM_DBL Then
                                                    g_arr_varHorses(i, 22) = g_arr_varHorses(i, 22) + 1
                                                Else
                                                    g_arr_varHorses(i, 22) = 1
                                                End If

                                            'For debugging purposes
                                            #If Debugging Then
                                                Debug.Print g_arr_varHorses(i, 1) & " runs in the slipstream of " _
                                                    & g_arr_varHorses(k, 1) & " (" & g_arr_varHorses(i, 22) & ")"
                                            #End If
                                            
                                        End If
                                End If

                            Next k
                        End If

                    'For debugging purposes
                    #If Debugging Then
                        If objOption.SLIPSTREAM Then Debug.Print g_arr_varHorses(i, 1) & " (Current position " _
                            & g_arr_varHorses(i, 4) & "): " & g_arr_varHorses(i, 8) & " (before slipstream effect)"
                    #End If

                    'Schrittweite pro Durchlauf festlegen (1 oder 2 Spalten)
                        g_arr_varHorses(i, 8) = g_arr_varHorses(i, 8) + g_arr_varHorses(i, 22) / 2000 'Windschatteneffekt addieren
                        g_arr_varHorses(i, 9) = Round(g_arr_varHorses(i, 8), 0) 'Runden auf ganze Zahlen (1 oder 2)
                        
                        g_arr_varHorses(i, 9) = g_arr_varHorses(i, 9) * objOption.SPEED_FACTOR 'take the race speed factor into account
                   
                    'For debugging purposes
                    #If Debugging Then
                        Debug.Print g_arr_varHorses(i, 1) & " (Current position " & g_arr_varHorses(i, 4) & "): " _
                            & g_arr_varHorses(i, 8) & " --> " & g_arr_varHorses(i, 9) & " steps"
                    #End If
                Next i 'Ende der Neuberechnung der Positionen der Pferde
                
            'Pferde laufen
                For i = 1 To UBound(g_arr_varHorses)
                    If g_arr_varHorses(i, 0) = "START" Then 'nur wenn Pferd am Start ist
                                            
                        'Pferd löschen
                            Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)) _
                                .Interior.color = objRace.TRACK_COLOUR
                                
                        'Neue Position des Pferds festlegen (nur wenn Pferd noch läuft)
                            If g_arr_varHorses(i, 0) = "START" Then
                                g_arr_varHorses(i, 4) = g_arr_varHorses(i, 4) + g_arr_varHorses(i, 9)
                                'RECORD.. g_arr_varHorses(i, 4) in Array schreiben, das ist neue die Position in dieser Runde
                                '   i ist die nummer im array(?)
                            End If
                            
                        'Windschatten zeichnen
                            If objOption.SLIPSTREAM And objOption.SLIPSTREAM_SHOW And g_arr_varHorses(i, 22) > 0 _
                                And g_arr_varHorses(i, 4) <= objRace.METERS + m_intLeftColumns + 7 Then
                                    Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9), _
                                        Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 12)).Interior.Pattern = xlGray16
                                'For debugging purposes
                                #If Debugging Then
                                    Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9).Value = "*"
                                #End If
                            End If
                            
                        'Wasserspritzer zeichnen
                            If objRace.SQUIRT = True And g_arr_varHorses(i, 4) > 13 + m_intLeftColumns _
                                And g_arr_varHorses(i, 4) <= objRace.METERS + m_intLeftColumns Then
                                
                                If Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)).Interior.color = objOption.PUDDLE_COL Then
                                
                                    'Calculate pattern by chance
    '                                    Randomize 'Zufallsgenerator zurücksetzen
                                        intSquirtPattern = Int((18 - 16 + 1) * Rnd + 16) ' 16=xlCrissCross 17=xlGray25 18=xlGray8
                                        intSquirtLength = Int((8 - 4 + 1) * Rnd + 4) 'value between 4 and 8
                                        dblSquirtColour = (Int(1.5 - 0 + 0) * Rnd + 0) - 1 'value between -1.0 and +1.0
    
                                        With Range(Cells(g_arr_varHorses(i, 3) - 1, g_arr_varHorses(i, 4) - 6), _
                                            Cells(g_arr_varHorses(i, 3) + 1, g_arr_varHorses(i, 4) - 6 - intSquirtLength)).Interior
                                                .Pattern = intSquirtPattern
                                                .PatternThemeColor = xlThemeColorDark1
                                                .PatternTintAndShade = Round(dblSquirtColour, 1)
                                        End With 'dblSquirtColour runden auf 1/10 sonst lauften wir in den Fehler >64000 versch. Zellformate
                                End If
                            End If
                            
                        'Pferd neu setzen (auch die, die schon im Ziel sind wegen Rendering)
                            Call PaintHorse(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7, g_arr_varHorses(i, 2))

                            'SPECIAL: In case of a mudflats race
                            If objRace.TRACK_SURFACE = "M" Then
                                For j = 0 To 7     'hide object under the horse
                                    If IsArray(g_arr_varHorses(i, 2)) Then
                                        Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7 + j) _
                                            .Font.color = g_arr_varHorses(i, 2)(j)
                                    Else
                                        Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7 + j) _
                                            .Font.color = g_arr_varHorses(i, 2)
                                    End If
                                Next j
                                Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8), _
                                    Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - 2 * objOption.SPEED_FACTOR)) _
                                        .Font.color = objOption.LUGWORM_COL
                                If Not IsEmpty(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8)) _
                                        And Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8) <> "|" Then
                                        With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8)
                                            .Value = ChrW(1154) 'lugworm stain
                                        End With
                                End If
                                
                                'Restore track colour behind the horse
                                For k = 1 To 2 * objOption.SPEED_FACTOR
                                    Select Case Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - k).Value
                                        Case "|"
                                            With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - k)
                                                .Font.color = objOption.PUDDLE_COL
                                                .Interior.color = objOption.PUDDLE_COL
                                            End With
                                            #If Debugging Then
                                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - k).Font.color = vbRed
                                            #End If
                                        Case Else
                                            With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8 - k)
                                                .Interior.color = objRace.TRACK_COLOUR
                                            End With
                                    End Select
                                Next k
                    
                            End If
                        
                        'SPECIAL: Wild boar devastation
                            If g_arr_varHorses(i, 24) = "WILD" _
                                 And g_arr_varHorses(i, 4) < objRace.METERS + m_intLeftColumns + 12 Then
                                 If IsArray(g_arr_varHorses(i, 2)) Then
                                    With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)
                                        .Font.color = g_arr_varHorses(i, 2)(0)
                                        .Value = "#"
                                    End With
                                Else
                                    With Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)
                                        .Font.color = g_arr_varHorses(i, 2)
                                        .Value = "#"
                                    End With
                                End If
                            End If

                        'Hufspuren zeichnen
                            If objOption.HOOFPRINTS And IsEmpty(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8)) _
                                And Not g_arr_varHorses(i, 24) = "WILD" Then _
                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8).Value = "'-"
                            
                    End If
                    
                    'Horizontal scrolling
                    If objOption.FOCUSED_RUN Then
                        'Focused Run
                        If g_arr_varHorses(i, 11) = objOption.FOCUSED_NR And g_arr_varHorses(i, 0) = "START" Then
                            If g_arr_varHorses(i, 4) > (ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).column _
                                                    - ActiveWindow.VisibleRange.column) / 2 Then 'Focused horse is in the middle of the screen
                                    ActiveWindow.ScrollColumn = ActiveWindow.ScrollColumn + g_arr_varHorses(i, 9)
                            End If
                        End If
                    Else
                        'No Focused Run
                        If g_arr_varHorses(i, 4) > ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).column _
                                                - ((ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).column _
                                                    - ActiveWindow.VisibleRange.column) * 1 / 10) _
                            And ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).column <= objRace.METERS + m_intLeftColumns + 2 Then
                                'Scrollen
                                ActiveWindow.ScrollColumn = ActiveWindow.VisibleRange.column _
                                                    + ((ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.count).column _
                                                    - ActiveWindow.VisibleRange.column) * 8 / 10)
                        End If
                    End If
                    
                    If objOption.RACE_INFO Then
                        'Get the position and name of the leader
                            If (g_arr_varHorses(i, 4) - m_intLeftColumns - 12) > m_intPosLeader Then
                                m_intPosLeader = g_arr_varHorses(i, 4) - m_intLeftColumns - 12
                                m_strNameLeader = g_arr_varHorses(i, 1)
                            End If
                        
                        'Refresh the race information data (in a pop-up)
                        If objOption.RACE_INFO_POP Then
                            'Display the race distance progress bar
                                If objOption.RACE_INFO_PROGRESS Then
                                    With frmRaceInfo.Controls("lbl_RI3b_dyn")
                                        .Width = CInt(200 / objRace.METERS * m_intPosLeader)
                                        .Caption = m_intPosLeader
                                    End With
                                End If
                            'Display the name of the leader
                                If objOption.RACE_INFO_LEADER Then
                                    If m_intPosLeader > 20 And m_intPosLeader < (objRace.METERS - 20) Then 'display name between 20m from the start to 20m to the finish
                                        frmRaceInfo.Controls("lbl_RI4a_dyn").Caption = strLeader
                                        frmRaceInfo.Controls("lbl_RI4b_dyn").Caption = m_strNameLeader
                                    Else
                                        frmRaceInfo.Controls("lbl_RI4a_dyn").Caption = " " 'for performance reasons
                                        frmRaceInfo.Controls("lbl_RI4b_dyn").Caption = " " 'for performance reasons
                                    End If
                                End If
                        End If
    
                        'Refresh the race information data (on the worksheet)
                        If objOption.RACE_INFO_WKS Then
                            'Display the race distance progress
                                If objOption.RACE_INFO_PROGRESS Then
                                    With g_wksRace.Cells(3 + m_intTopRows, 11)
                                        .Value = m_intPosLeader & strMeter & " / " & objRace.METERS & strMeter 'metres run
                                    End With
                                    If (intPipesProgressBar / objRace.METERS * m_intPosLeader) < 1 Then
                                        g_wksRace.Cells(3 + m_intTopRows, 12).Value = "|" 'for performance reasons
                                    Else
                                        With g_wksRace.Cells(3 + m_intTopRows, 12)
                                            .Value = String((intPipesProgressBar / objRace.METERS * m_intPosLeader), "|") 'progress bar
                                        End With
                                    End If
                                End If
                                
                            'Display the name of the leader
                                If objOption.RACE_INFO_LEADER Then
                                    If m_intPosLeader > 20 And m_intPosLeader < (objRace.METERS - 20) Then 'display name between 20m from the start to 20m to the finish
                                        With g_wksRace.Cells(1 + m_intTopRows, 2)
                                            .Value = strLeader
                                        End With
                                        With g_wksRace.Cells(2 + m_intTopRows, 10)
                                            .Value = m_strNameLeader
                                        End With
                                    Else
                                        g_wksRace.Cells(1 + m_intTopRows, 2).Value = " " 'for performance reasons
                                        g_wksRace.Cells(2 + m_intTopRows, 10).Value = " " 'for performance reasons
                                    End If
                                End If
                        End If
                        
                    End If
                Next i
                
            'Check ob ein Pferd im Ziel ist
                For i = 1 To UBound(g_arr_varHorses)
                    If g_arr_varHorses(i, 0) = "START" Then 'nur wenn Pferd noch läuft
                    
                    'For debugging purposes
                    #If Debugging Then
                        Debug.Print g_arr_varHorses(i, 1) & " - Position: " & g_arr_varHorses(i, 4)
                    #End If
                    
                        If g_arr_varHorses(i, 4) >= objRace.METERS + m_intLeftColumns + 12 Then 'Ziellinie erreicht
                        'Jump 3
                            g_arr_varHorses(i, 0) = "BERECHNUNG"
                            m_intHorsesFinishing = m_intHorsesFinishing + 1 'zählen, wie viele Pferde in diesem Durchlauf ins Ziel kommen
                        End If
                    End If
                Next i
                
                If m_intHorsesFinishing > 0 Then
                    If m_blnWin = False Then
                        If m_intHorsesFinishing > 1 Then
                            m_blnPhotofinish = True 'Fotofinish
                            'Text anpassen
                                g_wksRace.Cells(2 + m_intTopRows, objRace.METERS + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = GetText(g_arr_Text, "RACE025")  '"FOTOFINISH!"
                        End If
                        Call CreateFinishPhoto 'Zielfoto machen
                    End If
                    m_blnWin = True 'damit das Zielfoto nur 1x gemacht wird
                    Call CalculateRanking 'Absprung in die Platzberechnung, wenn mehr als ein Pferd in dieser Runde ins Ziel kommen
                End If
            DoEvents 'Rendering
        Loop

        'For debugging purposes: Calculate race time
        #If Debugging Then
            Debug.Print "Race finish: " & Format(Now, "HH:MM:SS")
            Debug.Print "Race time  : " & Format(Now - timeStart, "HH:MM:SS")
        #End If
        
    'Close the pop-up with race information
        If objOption.RACE_INFO Then
            If objOption.RACE_INFO_POP Then Unload frmRaceInfo 'close the pop-up
            If objOption.RACE_INFO_WKS Then Call basAuxiliary.RaceInfoWorksheet(xlNone, 0, m_intTopRows, False) 'reset: white cell, black font
        End If
    
    'In case of a photo finish
        If m_blnPhotofinish = True Then
            'Delay
                If Not DEVMODE Then Application.wait (Now + TimeValue("0:00:02"))
            'Unfreeze the window pane if it was frozen
                If objOption.NAMES_LEFT Or objOption.COLOURS_LEFT Or objOption.HIGHLIGHT_FAV _
                    Or (objOption.FOCUSED_RUN And objOption.HIGHLIGHT_FOC) Or (objOption.RACE_INFO And objOption.RACE_INFO_WKS) Then Call basAuxiliary.Freeze(0, 0, False)
            'Scroll
                On Error Resume Next
                ActiveWindow.ScrollColumn = objRace.METERS + m_intLeftColumns + 6 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)
                On Error GoTo 0
            'Black background
                g_wksRace.Range(Cells(5 + m_intTopRows, objRace.METERS + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
                    Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METERS + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))).Interior.ColorIndex = 1
            'Text
                g_wksRace.Cells(2 + m_intTopRows, objRace.METERS + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = ""
                g_wksRace.Cells(4 + m_intTopRows, objRace.METERS + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = GetText(g_arr_Text, "RACE026")  '("Photo creation")
            'Delay
                If Not DEVMODE Then Application.wait (Now + TimeValue("0:00:04"))
            'Show the finishing photo
                Call DrawFinishPhoto
            'Text
                g_wksRace.Cells(4 + m_intTopRows, objRace.METERS + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = ""
                g_wksRace.Cells(4 + m_intTopRows, objRace.METERS + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = GetText(g_arr_Text, "RACE027")  '("Photo evaluation")
            'Delay
                If Not DEVMODE Then Application.wait (Now + TimeValue("0:00:04"))
            'Text
                g_wksRace.Cells(4 + m_intTopRows, objRace.METERS + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = GetText(g_arr_Text, "RACE032")  '("Finishing photo")
        End If

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Platzierung berechnen wenn ein oder mehrere Pferde in einem Schleifendurchlauf ins Ziel kommen
Private Sub CalculateRanking()
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    ReDim m_arr_varResultsCalc(1 To m_intHorsesFinishing, 0 To 6)
    
    Dim intRandom As Integer 'Variable für Zufallszahlen
    Dim p As Boolean 'Platzierung wird vergeben wenn TRUE
    Dim intAssigned As Integer 'Berechnung intAssigned wenn alle Plätze vergeben sind
    Dim m As Integer 'Zählvariable

    m_intFinishLoop = m_intFinishLoop + 1 'Berechnungsrunde hochzählen
    p = False
    intAssigned = 0
    m = 1

    'Positionen in Berechnungs-Array eintragen
    For i = 1 To UBound(g_arr_varHorses)
        If g_arr_varHorses(i, 0) = "BERECHNUNG" Then
            g_arr_varHorses(i, 0) = "ZIEL" 'Finalen Status setzen
            m_arr_varResultsCalc(m, 1) = g_arr_varHorses(i, 11) 'Startnummer
            m_arr_varResultsCalc(m, 2) = g_arr_varHorses(i, 1) 'Name des Pferds
            m_arr_varResultsCalc(m, 3) = g_arr_varHorses(i, 4) 'Position des Pferds
            m_arr_varResultsCalc(m, 4) = g_arr_varHorses(i, 2) 'Farbe des Pferds
            m_arr_varResultsCalc(m, 5) = g_arr_varHorses(i, 23) 'Photo of the winner
            m_arr_varResultsCalc(m, 6) = g_arr_varHorses(i, 24) 'For special purposes
            m = m + 1
        End If
    Next i

    'Exakte Position des Pferds per Zufallszahlen generieren
    For i = 1 To UBound(m_arr_varResultsCalc)
        'Position durch Zufall neu berechnen
            Randomize 'Zufallsgenerator zurücksetzen
            intRandom = (Int((2 - 1 + 1) * Rnd + 1)) '1 = addieren, 2 = subtrahieren
            Randomize 'Zufallsgenerator zurücksetzen
            If intRandom = 1 Then
                m_arr_varResultsCalc(i, 3) = Round(m_arr_varResultsCalc(i, 3) _
                    + (Int((5 - 1 + 1) * Rnd + 1) / 10), 1) 'Dezimalstellen x.1 bis x.5
            Else
                m_arr_varResultsCalc(i, 3) = Round(m_arr_varResultsCalc(i, 3) _
                    - (Int((4 - 0 + 0) * Rnd + 0.5) / 10), 1) 'Dezimalstellen (x-1).6 bis x
            End If
    Next i
    
    'Platzierungen vergeben
    Do Until intAssigned >= UBound(m_arr_varResultsCalc)
        For i = 1 To UBound(m_arr_varResultsCalc)
            If m_arr_varResultsCalc(i, 0) <> "PLATZIERT" Then
                For j = i To UBound(m_arr_varResultsCalc)
                    If m_arr_varResultsCalc(j, 0) <> "PLATZIERT" Then
                        If m_arr_varResultsCalc(i, 3) >= m_arr_varResultsCalc(j, 3) Then 'Position ist größer als die des Vergleichspferds
                                p = True
                        Else 'Vergleichspferd ist weiter vorne
                            p = False
                            Exit For
                        End If
                    End If
                Next j
                If p = True Then
                    m_arr_varResultsCalc(i, 0) = "PLATZIERT" 'eintragen, dass das Pferd nicht mehr verglichen wird
                    intAssigned = intAssigned + 1 'hochzählen
                    p = False 'zurücksetzen
                    'Pferd in Ergebnisliste eintragen
                        m_arr_varResults(m_intPlace, 0) = m_intFinishLoop 'Berechnungsrunde
                        m_arr_varResults(m_intPlace, 2) = m_arr_varResultsCalc(i, 1) 'Startnummer
                        m_arr_varResults(m_intPlace, 3) = m_arr_varResultsCalc(i, 2) 'Name des Pferds
                        m_arr_varResults(m_intPlace, 4) = m_arr_varResultsCalc(i, 4) 'Farbe des Pferds
                        m_arr_varResults(m_intPlace, 5) = m_arr_varResultsCalc(i, 3) 'Position des Pferds
                        m_arr_varResults(m_intPlace, 6) = m_arr_varResultsCalc(i, 5) 'Photo of the winner
                        m_arr_varResults(m_intPlace, 7) = m_arr_varResultsCalc(i, 6) 'For special purposes
                        'Platzierung berechnen
                            If m_arr_varResults(m_intPlace, 0) = m_arr_varResults(m_intPlace - 1, 0) And _
                                m_arr_varResults(m_intPlace, 5) = m_arr_varResults(m_intPlace - 1, 5) Then
                                    'wenn exakt gleich wie das Pferd zuvor in dieser Berechnungsrunde
                                    m_arr_varResults(m_intPlace, 1) = m_arr_varResults(m_intPlace - 1, 1)
                            Else
                                'wenn Position kleiner ist als beim Pferd zuvor
                                m_arr_varResults(m_intPlace, 1) = m_intPlace
                            End If
                        m_intPlace = m_intPlace + 1 'Platz für nächstes Pferd hochzählen
                    Exit For
                End If
            End If
            
        Next i
    Loop
    
    'Exakte Position in Array für Zielfoto eintragen (Wert überschreiben)
    If m_blnExactFinish = False Then
        For i = 1 To UBound(m_arr_varResultsCalc)
            For j = 1 To UBound(m_arr_varPhotofinish)
                If m_arr_varResultsCalc(i, 1) = m_arr_varPhotofinish(j, 2) Then 'Abgleich der Startnummer
                    m_arr_varPhotofinish(j, 1) = m_arr_varResultsCalc(i, 3) 'Exakte Position in 1/10
                    m_arr_varPhotofinish(j, 3) = "X" 'Kennzeichen, dass die exakte Position schon berechnet wurde
                End If
            Next j
        Next i
        
        'Exakte Position der hinteren Pferde ermitteln und Wert überschreiben
        For i = 1 To UBound(m_arr_varPhotofinish)
            If m_arr_varPhotofinish(i, 3) <> "X" Then 'nur wenn noch nicht berechnet
                'Position durch Zufall neu berechnen
                Randomize 'Zufallsgenerator zurücksetzen
                intRandom = (Int((2 - 1 + 1) * Rnd + 1)) '1 = addieren, 2 = subtrahieren
                Randomize 'Zufallsgenerator zurücksetzen
                If intRandom = 1 Then
                    m_arr_varPhotofinish(i, 1) = Round(m_arr_varPhotofinish(i, 1) _
                        + (Int((5 - 1 + 1) * Rnd + 1) / 10), 1) 'Dezimalstellen x.1 bis x.5
                Else
                    m_arr_varPhotofinish(i, 1) = Round(m_arr_varPhotofinish(i, 1) _
                        - (Int((4 - 0 + 0) * Rnd + 0.5) / 10), 1) 'Dezimalstellen (x-1).6 bis x
                End If
                m_arr_varPhotofinish(i, 3) = "X" 'Kennzeichen, dass die exakte Position schon berechnet wurde
            End If
        Next i
        m_blnExactFinish = True
    End If
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Rank horses that did not finish
Private Sub NotFinished()
    For i = 1 To UBound(g_arr_varHorses)
        If g_arr_varHorses(i, 0) = "REFUSED" Then
            For j = 1 To UBound(m_arr_varResults) 'find the next line
                If m_arr_varResults(j, 1) = "" Then
                    m_arr_varResults(j, 1) = "-" 'ranking
                    m_arr_varResults(j, 2) = g_arr_varHorses(i, 11) 'horse number
                    m_arr_varResults(j, 3) = g_arr_varHorses(i, 1) 'horse name
                    m_arr_varResults(j, 4) = g_arr_varHorses(i, 2) 'horse colour
                    Exit For
                End If
            Next j
        End If
    Next i
End Sub

'Zielfoto machen wenn das erste Pferd das Ziel erreicht
Private Sub CreateFinishPhoto()
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
            
    'Daten eintragen
        For j = 1 To UBound(m_arr_varPhotofinish)
            m_arr_varPhotofinish(j, 0) = g_arr_varHorses(j, 3) 'Spur
            m_arr_varPhotofinish(j, 1) = g_arr_varHorses(j, 4) 'Position des Pferds
            m_arr_varPhotofinish(j, 2) = g_arr_varHorses(j, 11) 'Startnummer
            m_arr_varPhotofinish(j, 4) = g_arr_varHorses(j, 24) 'For special purposes
        Next j
    'Blitz wenn Fotofinish
        If m_blnPhotofinish Then
            For k = 1 To 8
                With g_wksRace.Range(Cells(5 + m_intTopRows, objRace.METERS + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
                    Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METERS + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)))
                        .Interior.ColorIndex = 1 'schwarz
                        .Interior.ColorIndex = 0 'weiß
                End With
            Next k
            g_wksRace.Range(Cells(5 + m_intTopRows, objRace.METERS + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
                Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METERS + m_intLeftColumns + 14 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))).Interior.color = objRace.TRACK_COLOUR
        End If

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Show the photo of the finish
Private Sub DrawFinishPhoto()
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    'Prepare a variable of type long, otherwise an overflow occurs in long distant races
    'when multiplying the track length for calculating the exact position
        Dim lngFin As Long
        lngFin = objRace.METERS 'copy the track length into the long variable
    
    'Draw the background
        Application.ScreenUpdating = False 'deactivate screen updating
        'Clear the cells
        g_wksRace.Range(Cells(5 + m_intTopRows, objRace.METERS + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METERS + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))).Clear
        'Photo frame
        g_wksRace.Range(Cells(5 + m_intTopRows, objRace.METERS + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METERS + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))) _
                .BorderAround ColorIndex:=0, Weight:=xlMedium
        If objOption.PHOTO_BW Then
            'Track and finish line black-and-white
            g_wksRace.Range(Cells(5 + m_intTopRows, objRace.METERS + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METERS + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))).Interior.ColorIndex = 1  'Rasen schwarz
            g_wksRace.Range(Cells(5 + m_intTopRows, objRace.METERS + m_intLeftColumns + 146 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METERS + m_intLeftColumns + 155 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))).Interior.ColorIndex = 0   'Zielline weiß
        Else
            'Track and finish line in original colours
            g_wksRace.Range(Cells(5 + m_intTopRows, objRace.METERS + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METERS + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))).Interior.color = objRace.TRACK_COLOUR
            g_wksRace.Range(Cells(5 + m_intTopRows, objRace.METERS + m_intLeftColumns + 146 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), Cells(objRace.NUMBER_ENROLLED * 2 + 7 + m_intTopRows, objRace.METERS + m_intLeftColumns + 155 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))).Interior.ColorIndex = 56   'Zielline grau
        End If
        
    'Draw the horses
        Dim fieldStart  As Integer, photoStart As Integer, horseHead As Integer, horseOffset As Integer, index As Integer
        fieldStart = objRace.METERS + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR) 'first column of painting field
        photoStart = objRace.METERS + m_intLeftColumns - 2

        For i = 1 To UBound(m_arr_varPhotofinish)
            If m_arr_varPhotofinish(i, 1) >= objRace.METERS + m_intLeftColumns - 2 Then 'only when the horse is in the photo
                'Jump 1
                'Draw the horse
                'Marco Frage am 3.1.19: Muss hier gerundet werden? Das ist doch immer eine ganze Zahl mit dem *10
                'Sollte immer eine ganze Zahl sein, Rundung stabilisiert zur Laufzeit und ist flexibler bei Aenderungen der Rennlogik
                horseHead = WorksheetFunction.Round((m_arr_varPhotofinish(i, 1) - photoStart) * 10, 0) - 1           'column of horses head
                horseOffset = 0                                     'distance between begin of the painting field and first part of the horse
                If horseHead > 79 Then horseOffset = horseHead - 79 'calculation of horseOffset: A hourse/pig is always 80 columns long
              
              'Marco Frage am 1.1.19: Kann die For-Schleife mit index = 1 anfangen und die If Not IsEmpty Abfrage raus?
              'Wenn auch zukuenftig immer sichergestellt ist, dass ab index 1 niemals empty sein wird. Sonst haben wir Laufzeitfehler. Da das Zielfoto nicht staendig neu gezeichnet wird habe ich den kleinen Performance-Verlust fuer stabileren Code hingenommen.
                For index = 0 To UBound(m_arr_varResults)
                    If Not IsEmpty(m_arr_varResults(index, 2)) Then
                        If m_arr_varResults(index, 2) = m_arr_varPhotofinish(i, 2) Then
                            Exit For
                        End If
                    End If
                Next index
                
                Dim pen As Integer, penMovement As Integer, horseSegment As Integer
                Dim horseSegmentColor As Long
                horseSegment = 8                      'the results show horses in 8 columns/segments
                pen = horseHead                       'start drawing at the head of the horse
                Do While pen > horseOffset
                    If pen - horseOffset >= 10 Then   'if possible draw 10 columns (1 segment) at once
                        penMovement = 10
                    Else
                        penMovement = pen - horseOffset                                     'if horse is not completely shown in the photo, draw until the photo end
                    End If
                    If IsArray(m_arr_varResults(index, 4)) Then
                        horseSegmentColor = m_arr_varResults(index, 4)(horseSegment - 1)    'get color from array
                    Else
                        horseSegmentColor = m_arr_varResults(index, 4)
                    End If
                    If objOption.PHOTO_BW Then horseSegmentColor = GreyToLong(CInt(RGBtoGrey(CLng(horseSegmentColor))))   'change to gray, if settings demand
                    Range(Cells(m_arr_varPhotofinish(i, 0), fieldStart + pen), _
                          Cells(m_arr_varPhotofinish(i, 0), fieldStart + pen - penMovement)) _
                          .Interior.color = horseSegmentColor
                    horseSegment = horseSegment - 1                                         'decrease segment counter to paint the next segment
                    pen = pen - penMovement                                                 'set pen to new position
                Loop
                
            End If

            'Horse names in the photo
            If objOption.NAMES_PHOTO = True And g_arr_varHorses(i, 0) <> "CANCELLED" Then
                With g_wksRace.Cells(g_arr_varHorses(i, 3), objRace.METERS + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))
                    .Font.name = "Arial"
                    .Font.Size = m_intFontSize
'                    .IndentLevel = 1
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                    If objOption.PHOTO_BW = True Then
                        .Font.color = vbWhite
                    Else
                        .Font.color = vbBlack
                    End If
                    .Value = g_arr_varHorses(i, 1)
                End With
            End If
        Next i

    'Activate screen updating
        Application.ScreenUpdating = True
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Race is finished - show info pop-up
Private Sub RaceFinished()
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    'Verzögerung
        If Not DEVMODE Then Application.wait (Now + TimeValue("0:00:02"))
    
    'Pop-up
        'Set the button mode
        g_strMsgButtons = "OK"
        'Assign the text for the pop-up
        g_strMsgCaption = objRace.RACE_NAME & " " & objRace.RACE_YEAR
        g_strMsgText = GetText(g_arr_Text, "RACE028") & vbNewLine & GetText(g_arr_Text, "RACE029")
        
        If objOption.SPEECH Then Call SpeechOut(g_strMsgText)
        
        'Display the pop-up
        frmMsg_MultiPurpose.show (vbModal) 'modal
        
    'Unfreeze the window pane
        If objOption.NAMES_LEFT Or objOption.COLOURS_LEFT Or objOption.HIGHLIGHT_FAV _
            Or (objOption.FOCUSED_RUN And objOption.HIGHLIGHT_FOC) Or (objOption.RACE_INFO And objOption.RACE_INFO_WKS) Then Call basAuxiliary.Freeze(0, 0, False)
    'Scrollen zu Ergebnistafel
        Call basAuxiliary.Scroll(objRace.METERS, m_intTopRows + (objRace.NUMBER_ENROLLED * 2 + 9))

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Show ranking list
Private Sub ShowRankingList(afterRace As Boolean)
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN

    'Anzeigetafel
        With g_wksRace.Range(Cells(objRace.NUMBER_ENROLLED * 2 + 20 + m_intTopRows, objRace.METERS + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)), _
            Cells(objRace.NUMBER_ENROLLED * 2 + 20 + m_intHorsesStarting + 1 + m_intTopRows, objRace.METERS + m_intLeftColumns + 175 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)))
                .Clear 'clear all cell values and formattings
                .BorderAround ColorIndex:=0, Weight:=xlMedium 'Rahmen
                .Interior.color = 16777215 'Hintergrund
                .Font.name = "Courier New"
                .Font.Size = 12
                .NumberFormat = "@" 'Textformat
        End With
        With g_wksRace.Cells(objRace.NUMBER_ENROLLED * 2 + 20 + m_intTopRows, objRace.METERS + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR))  'Überschrift
            .Font.Size = 14
            .Font.Bold = True 'Fettschrift
            .IndentLevel = 1 'Text eingerückt
        End With
    
    'Ergebnisse eintragen
        'Überschrift
            g_wksRace.Cells(objRace.NUMBER_ENROLLED * 2 + 20 + m_intTopRows, objRace.METERS + m_intLeftColumns + 16 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value _
                = objRace.RACE_NAME & " " & objRace.RACE_YEAR & " - " & objRace.TRACK_LOCATION
        'Verzögerung
            If Not DEVMODE And objOption.RANKING_DELAY And afterRace Then Application.wait (Now + TimeValue("0:00:04"))
        'Platzierungen anzeigen
            Dim intPositionName As Integer 'Position of the horse names
            intPositionName = 0
            For i = UBound(m_arr_varResults) To 1 Step -1
                m_arr_varResults(i, 8) = i
                If objOption.RANKING_COL Then
                    'Show colours
                    intPositionName = 12
                    Call PaintHorse(objRace.NUMBER_ENROLLED * 2 + 20 + i + m_intTopRows, objRace.METERS + m_intLeftColumns + 19 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR), m_arr_varResults(i, 4))
                End If
                Cells(objRace.NUMBER_ENROLLED * 2 + 20 + i + m_intTopRows, objRace.METERS + m_intLeftColumns + 22 + intPositionName + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = m_arr_varResults(i, 1) & "."  'Platzierung
                Cells(objRace.NUMBER_ENROLLED * 2 + 20 + i + m_intTopRows, objRace.METERS + m_intLeftColumns + 29 + intPositionName + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = m_arr_varResults(i, 3) & _
                    " (#" & m_arr_varResults(i, 2) & ")" 'Name und Startnummer des Pferds
                'Wenn Checkbox zur Spannungssteigerung angehakt ist
                If Not DEVMODE And objOption.RANKING_DELAY And afterRace Then Application.wait (Now + TimeValue("0:00:01")) 'Verzögerung
            Next i

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Kurzfristige Geschwindigkeit der Pferde (Faktor wird bei jedem Schleifendurchlauf neu berechnet)
Function SpeedLoop() As Double
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    Randomize 'Zufallsgenerator zurücksetzen
    SpeedLoop = (Int((m_lngSpeedLoopHigh - m_lngSpeedLoopLow + 1) * Rnd + m_lngSpeedLoopLow) + 100000) / 100000 'Zufallszahl
    
    Exit Function
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Function

'Pferd mit Siegerkranz zeichnen
Private Sub DrawWinnerPhoto()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
        
    'Name des Siegers
        m_strWinner = ""
        For i = 1 To UBound(m_arr_varResults)
            If m_arr_varResults(i, 1) = 1 Then
                If i > 1 Then 'wenn mehrere Pferde gewinnen
                    m_strWinner = m_strWinner & " und "
                    m_blnDeadHeat = True
                    If Not DEVMODE Then Application.wait (Now + TimeValue("0:00:02"))     'Verzögerung
                End If
                m_strWinner = m_strWinner & UCase(m_arr_varResults(i, 3))
                'Paint the horse
                    Call PaintPicture( _
                        g_wksPIC, _
                        m_arr_varResults(i, 6), _
                        objRace.METERS + m_intLeftColumns + 175 + 19 + 2 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR), _
                        m_intTopRows + objRace.NUMBER_ENROLLED * 2 + 23 + 12 + 5, _
                        m_intTopRows + objRace.NUMBER_ENROLLED * 2 + 23, _
                        objRace.METERS + m_intLeftColumns + 177 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR))

            End If
        Next i
        g_wksRace.Cells(objRace.NUMBER_ENROLLED * 2 + 20 + m_intTopRows, objRace.METERS + m_intLeftColumns + 177 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR)).Value = GetText(g_arr_Text, "RACE031")
        g_wksRace.Cells(objRace.NUMBER_ENROLLED * 2 + 21 + m_intTopRows, objRace.METERS + m_intLeftColumns + 179 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR)).Value = m_strWinner
 
        If objOption.SPEECH Then Call SpeechOut(GetText(g_arr_Text, "RACE031"))
        If objOption.SPEECH Then Call SpeechOut(m_strWinner)
        
    'In case of a dead heat (more than one winner)
        If m_blnDeadHeat Then
            'Pop-up
                'Set the button mode
                g_strMsgButtons = "OK"
                'Assign the text for the pop-up
                g_strMsgCaption = objRace.RACE_NAME & " " & objRace.RACE_YEAR
                g_strMsgText = " " & UCase(GetText(g_arr_Text, "RACE033")) & "!"
                'Display the pop-up
                frmMsg_Info.show (vbModal) 'modal
        End If
        
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Analyze bet slips
Private Sub UserFormAnalyseBetSlips()
    
    Dim id As String
    Dim nm As String
    Dim ty As String
    Dim st As Double
    Dim od As Double
    Dim bt() As Integer
    Dim payout As Boolean
    Dim noWinner As Boolean
    
    Dim totalStake As Double
    Dim totalPayout As Double
    
    If g_colBetSlips.count > 0 Then

        'Headline: "Official racing result"
        Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
        With g_objLabel
            With .Font
                .name = "Tahoma"
                .Size = 12
                .Bold = True
            End With
            .left = 15
            .top = 10
            .Width = 300
            .TextAlign = fmTextAlignLeft
            .Caption = GetText(g_arr_Text, "BET039") 'Official racing result
        End With

        'Race result (place 1-4)
        Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
        With g_objLabel
            .Font.name = "Tahoma"
            .Font.Size = 10
            .left = 40
            .top = 30
            .Width = 300
            .Height = 50
            .TextAlign = fmTextAlignLeft
            For i = 1 To 4 'compile the label with the horses on place 1-4
                .Caption = .Caption & GetText(g_arr_Text, "BET040") & " " & i & ": " _
                                & m_arr_varResults(i, 3) & " (#" & m_arr_varResults(i, 2) & ")" & vbNewLine
            Next i
        End With

        'Placed bets
            'Headline
            Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
            With g_objLabel
                With .Font
                    .name = "Tahoma"
                    .Size = 12
                    .Bold = True
                End With
                .left = 15
                .top = 90
                .Width = 300
                .TextAlign = fmTextAlignLeft
                .Caption = GetText(g_arr_Text, "BET042") 'Placed bets
            End With
    
    noWinner = True 'Reset variable
    k = 1 'lfd. nr. wettschein
             
        For i = 1 To g_colBetSlips.count
            payout = True
            id = g_colBetSlips(i).id
            nm = g_colBetSlips(i).GamblerName
            ty = g_colBetSlips(i).BType
            st = g_colBetSlips(i).Stake
            od = g_colBetSlips(i).Odd * 10
            bt() = g_colBetSlips(i).bet
            
            'Analyse bet slips
            Dim payCash As String
            Dim payColor As Long
            For j = 1 To UBound(bt)
                If bt(j) <> m_arr_varResults(j, 2) Then payout = False
            Next j
            If payout = False Then
                payCash = 0
                payColor = &H8080FF 'red ...was ist das für ein Format?? --> Buch
            Else
                payCash = st / 10 * od
                payColor = 52377 'green
            End If
            
            'Write data of each bet slip to the userform
                'Name and bet slip ID
                Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
                With g_objLabel
                    With .Font
                        .name = "Tahoma"
                        .Size = 10
                        .Bold = True
                    End With
                    .left = 40
                    .top = 98 + k * 12
                    .Width = 350
                    .TextAlign = fmTextAlignLeft
                    .Caption = nm & " (" & GetText(g_arr_Text, "BET001") & " " & GetText(g_arr_Text, "ODD001") & " " & id & ")"
                End With
                
                k = k + 1
                
                'Type of bet
                Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
                With g_objLabel
                    .Font.name = "Tahoma"
                    .Font.Size = 10
                    .left = 80
                    .top = 98 + k * 12
                    .Width = 200
                    .TextAlign = fmTextAlignLeft
                    .Caption = UCase(GetText(g_arr_Text, "BET007")) & ": " & ty 'TYPE OF BET: xxxxx
                End With
            
                k = k + 1
                
            For j = 1 To UBound(bt)
                'Guess
                Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
                With g_objLabel
                    .Font.name = "Tahoma"
                    .Font.Size = 10
                    .left = 80
                    .top = 98 + k * 12
                    .Width = 200
                    .TextAlign = fmTextAlignLeft
                    .Caption = GetText(g_arr_Text, "BET041") & ": " & GetHorseName(bt(j)) & " (#" & bt(j) & ")"
                End With
                
                k = k + 1
                
            Next j
            
            'Stake
                Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
                With g_objLabel
                    .Font.name = "Tahoma"
                    .Font.Size = 10
                    .left = 80
                    .top = 98 + k * 12
                    .Width = 100
                    .TextAlign = fmTextAlignLeft
                    .Caption = GetText(g_arr_Text, "BET037") & ": " & Format(st, "0.00") & " " & GetText(g_arr_Text, "BET035") 'Stake: xx EUR
                End With
                
            'Pay-out
                Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
                With g_objLabel
                    .Font.name = "Tahoma"
                    .Font.Size = 10
                    .left = 200
                    .top = 98 + k * 12
                    .Width = 150
                    .TextAlign = fmTextAlignLeft
                    .Caption = "  " & GetText(g_arr_Text, "BET038") & ": " & Format(payCash, "0.00") & " " & GetText(g_arr_Text, "BET035") 'Pay-out: xx EUR
                    .BackColor = payColor
                End With
                
                'For statistic purposes
                    totalStake = totalStake + st
                    totalPayout = totalPayout + payCash
                
                k = k + 2
            
        Next i

            'Number of bet slips
            Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
            With g_objLabel
                With .Font
                    .name = "Tahoma"
                    .Size = 10
                End With
                .left = 40
                .top = 98 + k * 12
                .Width = 300
                .TextAlign = fmTextAlignLeft
                .Caption = GetText(g_arr_Text, "START012") & ": " & g_colBetSlips.count 'Number of bet slips
            End With
            
            k = k + 1
            
            'Total stake / payout
            Set g_objLabel = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
            With g_objLabel
                With .Font
                    .name = "Tahoma"
                    .Size = 10
                End With
                .left = 40
                .top = 98 + k * 12
                .Width = 300
                .TextAlign = fmTextAlignLeft
                .Caption = GetText(g_arr_Text, "BET043") & ": " & totalStake & " " & GetText(g_arr_Text, "BET035") & " / " _
                            & GetText(g_arr_Text, "BET044") & ": " & totalPayout & " " & GetText(g_arr_Text, "BET035")
            End With

        With frmBettingAnalysis
            .Caption = objRace.RACE_NAME & " " & objRace.RACE_YEAR & " | " & objRace.TRACK_NAME & ", " & objRace.TRACK_LOCATION _
                        & " (" & objRace.COUNTRY & ")"
            .Width = 400 'width of the pop-up

            If g_colBetSlips.count <= 5 Then
                .Height = 98 + k * 12 + 50 'height of the pop-up depending of the number of bets placed
            Else
                .Height = 440 'fix height if more than 5 bets are placed
            End If
            .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
            .ScrollHeight = 98 + k * 12 + 30 'height of the vertical scrolling
            .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
            'Position of the pop-up
                .StartUpPosition = 0
                .top = ActiveWindow.top + ((ActiveWindow.Height - .Height) / 2) 'vertically centred
                .left = ActiveWindow.left + ((ActiveWindow.Width - .Width) - ActiveWindow.Width / 10) 'near the right border
            .show (vbModeless)  'modeless
        End With
     
    End If
End Sub

'Retrieve horse name from the horse number
Private Function GetHorseName(num As Integer) As String
    Dim x As Integer
    For x = 1 To UBound(g_arr_varHorses())
        If num = g_arr_varHorses(x, 11) Then Exit For
    Next x
    GetHorseName = g_arr_varHorses(x, 1)
End Function

'Warnhinweis
Private Sub ShowWarning()
    Dim strWarningMessage As String
    
    'In case an error occurs
    On Error GoTo ERRORHANDLING 'REIN
    
    'Compose the warning message
    strWarningMessage = strWarningMessage & GetText(g_arr_Text, "WARN001") & vbNewLine & GetText(g_arr_Text, "WARN002")

    'Pop-up
        'Set the button mode
        g_strMsgButtons = "OK"
        'Assign the text for the pop-up
        g_strMsgCaption = GetText(g_arr_Text, "USERFORM003")
        g_strMsgText = strWarningMessage
        'Display the pop-up
        frmMsg_Attention.show (vbModal) 'modal
        
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Delete all GaloppSim worksheets
Private Sub AI_DeleteWorksheets()
    On Error Resume Next
        Application.DisplayAlerts = False 'Turn off application warnings
        'Delete worksheets
            g_wksRace.Delete
        Application.DisplayAlerts = True 'Turn on application warnings
    On Error GoTo 0
End Sub

'Information for UserForm "frmStart"
Private Sub UserFormSTART()
    With frmStart
        .Caption = g_c_tool
        .lblS1.Caption = objRace.RACE_NAME & " " & objRace.RACE_YEAR 'race name and year
        .lblS2.Caption = objRace.RACE_TYPE_TEXT & " " & GetText(g_arr_Text, "RACE007") & " " & objRace.METERS & " " & GetText(g_arr_Text, "RACE009") ' race type and distance
        .lblS3.Caption = objRace.TRACK_NAME & " " & GetText(g_arr_Text, "RACE002") & " " & objRace.TRACK_LOCATION & " (" & objRace.COUNTRY & ")" 'race course, loaction and country
        .lblS6.Caption = m_intHorsesStarting & " " & g_arr_Grammar(4) 'number of horses starting
        If objRace.REAL_RACE = "Y" Then
            .lblS10.Caption = UCase(GetText(g_arr_Text, "START009")) 'REAL RACE
        Else
            .lblS10.Caption = UCase(GetText(g_arr_Text, "START010")) 'FICTITIOUS RACE
        End If
        .lblS8.Caption = GetText(g_arr_Text, "START005") 'caption of the betting section
        Call NumberBetSlips 'refresh the number of bet slips
        .lblFocus.Caption = g_arr_Grammar(1) & " " & GetText(g_arr_Text, "START006") 'Label "Horse in focus"
        .cmdS1.Caption = GetText(g_arr_Text, "START002") 'button "Add bet slip"
        .cmdS2.Caption = GetText(g_arr_Text, "START003") 'button "Start race"
        With .lblS4 'track surface
            .Caption = objRace.TRACK_SURFACE_TEXT 'text with the surface
            .BorderStyle = fmBorderStyleSingle 'draw a border around
            .BackColor = objRace.TRACK_COLOUR 'set the color according to the track colour
        End With
        .cmdS5.Caption = GetText(g_arr_Text, "START014") 'button "Show speed and form"
        .cmdS6.Caption = GetText(g_arr_Text, "START015") 'button "Show fixed odds"
        If objOption.BET_MODE = True And objRace.BETTING_ALLOWED = "Y" Then 'height of the pop-up
            .Height = 315 'if the betting mode is enabled
        Else
            .Height = 200 'if the betting mode is disabled
        End If
        .show (vbModal) 'modal
    End With
End Sub

Public Sub NumberBetSlips()
    frmStart.lblBet02.Caption = GetText(g_arr_Text, "START012") & ": " & g_colBetSlips.count 'number of bet slips filed in
End Sub

'Name of the gambler for placing a bet
Public Sub Gambler()
    
    'Set the button mode
    g_strMsgButtons = "CancelOK"
    'Assign the text for the pop-up
    g_strMsgText = GetText(g_arr_Text, "BET002")
    'Display the pop-up
    With frmInp_MultiPurpose
        .Caption = objRace.RACE_NAME & " " & objRace.RACE_YEAR
        .show (vbModal) 'modal
    End With
    'Evaluate the name of the player
    If g_strButtonPressed = "OK" And Trim(g_strPlayerName) <> "" Then _
        Call UserFormBetSlip(g_strPlayerName)

End Sub

'Information for UserForm "BetSlip"
Private Sub UserFormBetSlip(strName As String)
    With frmBetSlip
        .Caption = strName
        .lblC1 = objRace.TRACK_NAME & " - " & objRace.TRACK_LOCATION & " (" & objRace.COUNTRY & ")"
        .lblC2 = objRace.RACE_NAME & " " & objRace.RACE_YEAR
        .show (vbModal) 'modal
    End With
End Sub

'Call UserForm with speed and odds
Public Sub ShowSpeed(ODDS As Boolean)
    Call UserFormOdds(ODDS)
End Sub

'Information for UserForm "Odds"
Private Sub UserFormOdds(ODDS As Boolean)
    
    Dim min As Integer, max As Integer
    Dim i As Integer, j As Integer, k As Integer
    
    With frmOdds
        .Caption = objRace.RACE_NAME & " " & objRace.RACE_YEAR
        .Width = 560
        .Height = 80
        .lblO0.Caption = GetText(g_arr_Text, "ODD001")
        .lblO1.Caption = GetText(g_arr_Text, "ODD002")
        With .lblO2
            .Caption = GetText(g_arr_Text, "ODD003")
            .ControlTipText = GetText(g_arr_Text, "ODD006")
            .TextAlign = fmTextAlignRight
        End With
        If Not ODDS Then .lblO2.Caption = "" 'delete caption "Odds"
    End With
    
    k = 1

    For i = 1 To UBound(g_arr_varHorses)
        If min = 0 Or g_arr_varHorses(i, 17) < min Then min = g_arr_varHorses(i, 17)
        If g_arr_varHorses(i, 17) > max Then max = g_arr_varHorses(i, 17)
    Next i
    
    For i = min To max
        For j = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(j, 17) = i Then
            
                Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True)
                With g_objLabel 'Nr and name of the horse
                    .Font.name = "Tahoma"
                    .Font.Size = 12
                    .left = 12
                    .top = 28 + k * 18
                    .Width = 200
                    .TextAlign = fmTextAlignLeft
                    .Caption = "#" & g_arr_varHorses(j, 11) & vbTab & g_arr_varHorses(j, 1)
                    If g_arr_varHorses(j, 0) <> "START" Then .Font.Strikethrough = True
                End With
                
                'Adjust UserForm height
                frmOdds.Height = frmOdds.Height + g_objLabel.Height
                
                If ODDS Then
                    Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True)
                    With g_objLabel 'odd
                        .Font.name = "Tahoma"
                        .Font.Size = 12
                        .left = 220
                        .top = 28 + k * 18
                        .Width = 62
                        .TextAlign = fmTextAlignRight
                        .Caption = g_arr_varHorses(j, 17) & ":10"
                        If g_arr_varHorses(j, 0) <> "START" Then .Font.Strikethrough = True
                    End With
                End If

                If g_arr_varHorses(j, 0) = "START" Then
                    Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True) 'upper horizontal bar
                    With g_objLabel '~speed
                        .left = 290
                        .top = 27 + k * 18
                        .Height = 7
                        If objRace.REAL_RACE = "Y" Then
                            .Width = 100 + (((g_arr_varHorses(j, 5) * 1000000) - 1499900) / 2)
                        Else
                            .Width = 100 + (((g_arr_varHorses(j, 6) * 1000000) - 1499900) / 2)
                        End If
                        .BackColor = 14395790 'blue
                    End With
                End If
                
                If g_arr_varHorses(j, 0) = "START" Then
                    Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True) 'lower horizontal bar
                    With g_objLabel '~condition
                        .left = 290
                        .top = 27 + k * 18 + 8
                        .Height = 7
                        .Width = 100 + (((g_arr_varHorses(j, 6) * 1000000) - 1499900) / 2) + g_arr_varHorses(j, 18)
                        .BackColor = 6740479 'yellow
''For development purposes only
'                        .Caption = "T: " & g_arr_varHorses(j, 6) * 1000000 & " ... " & g_arr_varHorses(j, 18)
                    End With
                End If

''For development purposes only
'                Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True)
'                With g_objLabel 'real condition
'                    .Font.Name = "Tahoma"
'                    .Font.Size = 12
'                    .Left = 580
'                    .Top = 13 + k * 18
'                    .Height = 15
'                    .Width = 100 + (((g_arr_varHorses(j, 6) * 1000000) - 1499900) / 2)
'                    .TextAlign = fmTextAlignLeft
'                    .BackColor = 6740479
'                    .Caption = "T: " & g_arr_varHorses(j, 6) * 1000000
'                End With
                
                k = k + 1
            End If
        Next j
    Next i

    Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True) 'upper horizontal bar
    With g_objLabel '~speed
        With .Font
                .name = "Tahoma"
                .Bold = True
                .Size = 10
        End With
        .left = 290
        .top = 6
        .Height = 15
        .Width = 246
        .TextAlign = fmTextAlignCenter
        .BackColor = 14395790 'blue
        .Caption = GetText(g_arr_Text, "ODD005")
        .ControlTipText = GetText(g_arr_Text, "ODD008") & " " & g_arr_Grammar(5)
    End With

    Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True) 'lower horizontal bar
    With g_objLabel '~condition
        With .Font
                .name = "Tahoma"
                .Bold = True
                .Size = 10
        End With
        .left = 290
        .top = 22
        .Height = 15
        .Width = 246
        .TextAlign = fmTextAlignCenter
        .BackColor = 6740479 'yellow
        .Caption = GetText(g_arr_Text, "ODD004")
        .ControlTipText = GetText(g_arr_Text, "ODD007") & " " & g_arr_Grammar(5)
    End With

    frmOdds.show (vbModal) 'modal
End Sub

'Call UserForm for showing receipt
Public Sub ShowReceipt(id As Integer)
    Call UserFormReceipt(id)
End Sub

'Information for UserForm "Receipt"
Private Sub UserFormReceipt(id As Integer)
    Dim bet As String, horsename As String
    Dim i As Integer, j As Integer
    Dim bt() As Integer
    bt() = g_colBetSlips(id).bet
    For i = 1 To UBound(bt)
        For j = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(j, 11) = bt(i) Then
                horsename = g_arr_varHorses(j, 1)
                Exit For
            End If
        Next j
        bet = bet & bt(i) & " " & horsename & vbNewLine
    Next i

    With frmReceipt
        .Caption = g_colBetSlips(id).GamblerName
        .lblR1 = UCase(objRace.TRACK_LOCATION) ' & " (" & objRace.COUNTRY & ")")
        .lblR2 = UCase(objRace.RACE_NAME)
        .lblR3 = m_intHorsesStarting & " " & UCase(g_arr_Grammar(4))
        .lblR4 = UCase(g_colBetSlips(id).BType)
        .lblR5 = UCase(bet)
        .lblR6 = UCase(GetText(g_arr_Text, "BET036") & " " & Format(g_colBetSlips(id).Stake, "0.00") & " " & GetText(g_arr_Text, "BET035"))
        .lblR7 = g_colBetSlips(id).id
        .show (vbModal) 'modal
    End With
End Sub


'CALLBACKS (Excel ribbon events)
'-------------------------------

'Callback for customUI.onLoad
Private Sub AI_GaloppSimAddinInitialize(ribbon As IRibbonUI)
    Set g_RibbonGaloppSim = ribbon
    If g_wksTEXT Is Nothing Then Set g_wksTEXT = tableTEXT
    Call GetTextComponents
End Sub

'Callback for lblDE getLabel
Private Sub AI_GetLabel(control As IRibbonControl, ByRef returnedVal)
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

'Callbacks for Tooltips
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

'Callbacks for button status
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

'Initialwerte der Checkboxen im Menüband (getPressed)
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

'Checkboxen im Menüband (onAction)
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

'Ribbon button "Race options"
Private Sub AI_OptionsRace(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    Call AnimalGrammar
    frmOptionsRace.show (vbModal) 'display UserForm (modal)
End Sub

'Startbutton im Menüband
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
                .Cells(basAuxiliary.GetRow(ThisWorkbook.Worksheets(objRace.SELECTED), "DISTANCE METERS"), 2).Value & "m) - " & _
                .Cells(basAuxiliary.GetRow(ThisWorkbook.Worksheets(objRace.SELECTED), "TRACK LOCATION"), 2).Value
        End With
        
'        'A MessageBox cannot handle unicode, so for example Cyrillic characters are displayed as question marks
'        If MsgBox((GetText(g_arr_Text, "RACE003") & " " & g_wksRace.OLEObjects("CBraces").Object.text), _
'            vbOKCancel, g_c_TOOL) = vbCancel Then Exit Sub

        'Pop-up
            'Set the button mode
            g_strMsgButtons = "CancelOK"
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_tool
            g_strMsgText = GetText(g_arr_Text, "RACE003") & ": " & strNextRace
            'Display the pop-up
            frmMsg_MultiPurpose.show (vbModal) 'modal
            'Evaluate the return value
            If g_strButtonPressed = "CANCEL" Then Exit Sub
            
            If g_strPlayMode = "AI" Then g_RibbonGaloppSim.Invalidate 'reset Excel ribbon
            Call ShowNewRaceScreen
    End If

    Call NewRace
End Sub

'Ergebnis-Button im Menüband
Private Sub AI_Results(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call RankingList
End Sub

'Ribbon button "Winner"
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

Private Sub ShowWinnerPhoto()
    If objRace.STARTED Then
        g_wksRace.Activate
        Call DrawWinnerPhoto
        Call basAuxiliary.Scroll(objRace.METERS + m_intLeftColumns + 9 + 160 + m_intColumsAfterFinish + (2 * 10 * objOption.SPEED_FACTOR), m_intTopRows + (objRace.NUMBER_ENROLLED * 2 + 9))
        If g_strPlayMode = "RS" Then frmRS_navigation.show (vbModeless) 'modeless
    End If
End Sub

Private Sub RankingList()
    If objRace.STARTED Then
        g_wksRace.Activate
        Call basAuxiliary.Scroll(objRace.METERS + m_intLeftColumns + 8 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR), m_intTopRows + (objRace.NUMBER_ENROLLED * 2 + 9))
        Call ShowRankingList(False)
        If g_strPlayMode = "RS" Then frmRS_navigation.show (vbModeless) 'modeless
    End If
End Sub

'Zielfoto-Button im Menüband
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

Private Sub ShowFinishPhoto()
    If objRace.STARTED Then
        g_wksRace.Activate
        'Text anpassen
        g_wksRace.Cells(4 + m_intTopRows, objRace.METERS + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR)).Value = GetText(g_arr_Text, "RACE032") '("Zielfoto")
        Call basAuxiliary.Scroll(objRace.METERS + m_intLeftColumns + 8 + m_intColumsAfterFinish + (2 * objOption.SPEED_FACTOR), m_intTopRows + 1)
        Call DrawFinishPhoto
        If g_strPlayMode = "RS" Then frmRS_navigation.show (vbModeless) 'modeless
    End If
End Sub

'Wett-Button im Menüband
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

Private Sub ShowBets()
    If objRace.STARTED And objOption.BET_PLACED Then
        Call UserFormAnalyseBetSlips
    End If
End Sub

'Info-Button im Menüband
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

'AI edition: Adapt Excel settings according to the selected mode
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

Public Sub AI_ExcelModeEnd()
    Select Case objOption.EXCEL_MODE
        Case "normal"
            
        Case "TVmenu"
            
        Case "TVfull"
            Call ExcelOptionsTVmenu
    End Select
End Sub

Private Sub ShowInfo()
    
    With frmInfo
        .Caption = g_c_tool & " - " & GetText(g_arr_Text, "INFO08")
        .lbl_info01.Caption = GetText(g_arr_Text, "GEN01") & vbNewLine & GetText(g_arr_Text, "GEN02")
        
        'MultiPage captions
            'MultiPage "Info"
            For i = 0 To 6
                .multiPage_info.Pages(i).Caption = GetText(g_arr_Text, "PAGE0" & i + 1)
            Next i
                .multiPage_info.Value = 0 'set the focus on the first page
            'MultiPage "Algorithms"
            For i = 0 To 8
                .multiPage_algo.Pages(i).Caption = GetText(g_arr_Text, "PAGEALGO0" & i + 1)
            Next i
                .multiPage_algo.MultiRow = True 'display tabs in rows without scrolling
                .multiPage_algo.Value = 0 'set the focus on the first page
        
        'Page "GaloppSim"
            With .lbl_info_galoppsim01
                .Caption = GetText(g_arr_Text, "INFO01") & vbNewLine & vbNewLine _
                            & GetText(g_arr_Text, "INFO02") & vbNewLine & vbNewLine _
                            & GetText(g_arr_Text, "INFO03") & vbNewLine & vbNewLine _
                            & GetText(g_arr_Text, "INFO04") & vbNewLine & vbNewLine _
                            & GetText(g_arr_Text, "INFO05") & vbNewLine & vbNewLine _
                            & GetText(g_arr_Text, "INFO06") & vbNewLine & vbNewLine _
                            & GetText(g_arr_Text, "INFO07") & vbNewLine & vbNewLine
                .Width = 460 'label width fix
                .AutoSize = True 'label height depending on the content
            End With
            With .multiPage_info.Pages(0)
                .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                .ScrollHeight = .lbl_info_galoppsim01.Height 'height of the vertical scrolling
                .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
            End With

        'Page "Development team"
            'Marco Matjes
            .lbl_info_team01a.Caption = GetText(g_arr_Text, "TEAM001")
            .lbl_info_team01b.Caption = GetText(g_arr_Text, "TEAM002")
            .img_info_team01.ControlTipText = GetText(g_arr_Text, "TEAM003")
            'Florian
            .lbl_info_team02a.Caption = GetText(g_arr_Text, "TEAM004")
            .lbl_info_team02b.Caption = GetText(g_arr_Text, "TEAM005")
            .img_info_team02.ControlTipText = GetText(g_arr_Text, "TEAM006")
            'Paul
            .lbl_info_team03a.Caption = GetText(g_arr_Text, "TEAM007")
            .lbl_info_team03b.Caption = GetText(g_arr_Text, "TEAM008")
            .img_info_team03.ControlTipText = GetText(g_arr_Text, "TEAM009")
            'Michael
            .lbl_info_team04a.Caption = GetText(g_arr_Text, "TEAM010")
            .lbl_info_team04b.Caption = GetText(g_arr_Text, "TEAM011")
            .img_info_team04.ControlTipText = GetText(g_arr_Text, "TEAM012")
            'Meike
            .lbl_info_team05a.Caption = GetText(g_arr_Text, "TEAM013")
            .lbl_info_team05b.Caption = GetText(g_arr_Text, "TEAM014")
            .img_info_team05.ControlTipText = GetText(g_arr_Text, "TEAM016")
            'Natalie
            .lbl_info_team06a.Caption = GetText(g_arr_Text, "TEAM016")
            .lbl_info_team06b.Caption = GetText(g_arr_Text, "TEAM017")
            .img_info_team06.ControlTipText = GetText(g_arr_Text, "TEAM018")
            'Atanas
            .lbl_info_team07a.Caption = GetText(g_arr_Text, "TEAM019")
            .lbl_info_team07b.Caption = GetText(g_arr_Text, "TEAM020")
            .img_info_team07.ControlTipText = GetText(g_arr_Text, "TEAM021")
            'Duncan
            .lbl_info_team08a.Caption = GetText(g_arr_Text, "TEAM022")
            .lbl_info_team08b.Caption = GetText(g_arr_Text, "TEAM023")
            .img_info_team08.ControlTipText = GetText(g_arr_Text, "TEAM024")
            
            'Vertical scrollbar
            With .multiPage_info.Pages(1)
                .ScrollBars = fmScrollBarsVertical
                .ScrollHeight = 400
                .KeepScrollBarsVisible = fmScrollBarsNone
            End With
            
        'Page "Algorithms"
            .img_info_algorithms01.ControlTipText = GetText(g_arr_Text, "ALGO01")
            .img_info_algorithms02.ControlTipText = GetText(g_arr_Text, "ALGO01")
            With .chk_info_algorithms01
                .Caption = GetText(g_arr_Text, "ALGO02")
                .Font.Size = 20
                .Font.Bold = True
                .ControlTipText = GetText(g_arr_Text, "ALGO03")
            End With
            'Algorithm 01 'overall race algorithm
                With .lbl_algo_01_00
                    .Caption = GetText(g_arr_Text, "PAGEALGO01")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_01_01
                    .Caption = GetText(g_arr_Text, "ALGO10")
                    .top = frmInfo.lbl_algo_01_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_01_02
                    .Caption = GetText(g_arr_Text, "ALGO11")
                    .top = frmInfo.lbl_algo_01_00.Height + frmInfo.lbl_algo_01_01.Height + 24
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .multiPage_algo(0)
                    .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                    .ScrollHeight = .lbl_algo_01_00.Height + .lbl_algo_01_01.Height + .lbl_algo_01_02.Height + 30 'height of the vertical scrolling
                    .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
                End With
            'Algorithm 02 '
                With .lbl_algo_02_00
                    .Caption = GetText(g_arr_Text, "PAGEALGO02")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_02_01
                    .Caption = GetText(g_arr_Text, "ALGO15")
                    .top = frmInfo.lbl_algo_02_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_02_02
                    .Caption = GetText(g_arr_Text, "ALGO16")
                    .top = frmInfo.lbl_algo_02_00.Height + frmInfo.lbl_algo_02_01.Height + 24
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .multiPage_algo(1)
                    .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                    .ScrollHeight = .lbl_algo_02_00.Height + .lbl_algo_02_01.Height + .lbl_algo_02_02.Height + 30 'height of the vertical scrolling
                    .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
                End With
            'Algorithm 03 '
                With .lbl_algo_03_00
                    .Caption = GetText(g_arr_Text, "PAGEALGO03")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_03_01
                    .Caption = GetText(g_arr_Text, "ALGO20")
                    .top = frmInfo.lbl_algo_03_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_03_02
                    .Caption = GetText(g_arr_Text, "ALGO21")
                    .top = frmInfo.lbl_algo_03_00.Height + frmInfo.lbl_algo_03_01.Height + 24
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .multiPage_algo(2)
                    .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                    .ScrollHeight = .lbl_algo_03_00.Height + .lbl_algo_03_01.Height + .lbl_algo_03_02.Height + 30 'height of the vertical scrolling
                    .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
                End With
            'Algorithm 04 '
                With .lbl_algo_04_00
                    .Caption = GetText(g_arr_Text, "PAGEALGO04")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_04_01
                    .Caption = GetText(g_arr_Text, "ALGO25")
                    .top = frmInfo.lbl_algo_04_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_04_02
                    .Caption = GetText(g_arr_Text, "ALGO26")
                    .top = frmInfo.lbl_algo_04_00.Height + frmInfo.lbl_algo_04_01.Height + 24
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .multiPage_algo(3)
                    .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                    .ScrollHeight = .lbl_algo_04_00.Height + .lbl_algo_04_01.Height + .lbl_algo_04_02.Height + 30 'height of the vertical scrolling
                    .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
                End With
            'Algorithm 05 '
                With .lbl_algo_05_00
                    .Caption = GetText(g_arr_Text, "PAGEALGO05")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_05_01
                    .Caption = GetText(g_arr_Text, "ALGO30")
                    .top = frmInfo.lbl_algo_05_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_05_02
                    .Caption = GetText(g_arr_Text, "ALGO31")
                    .top = frmInfo.lbl_algo_05_00.Height + frmInfo.lbl_algo_05_01.Height + 24
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .multiPage_algo(4)
                    .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                    .ScrollHeight = .lbl_algo_05_00.Height + .lbl_algo_05_01.Height + .lbl_algo_05_02.Height + 30 'height of the vertical scrolling
                    .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
                End With
            'Algorithm 06 '
                With .lbl_algo_06_00
                    .Caption = GetText(g_arr_Text, "PAGEALGO06")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_06_01
                    .Caption = GetText(g_arr_Text, "ALGO35")
                    .top = frmInfo.lbl_algo_06_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_06_02
                    .Caption = GetText(g_arr_Text, "ALGO36")
                    .top = frmInfo.lbl_algo_06_00.Height + frmInfo.lbl_algo_06_01.Height + 24
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .multiPage_algo(5)
                    .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                    .ScrollHeight = .lbl_algo_06_00.Height + .lbl_algo_06_01.Height + .lbl_algo_06_02.Height + 30 'height of the vertical scrolling
                    .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
                End With
            'Algorithm 07 '
                With .lbl_algo_07_00
                    .Caption = GetText(g_arr_Text, "PAGEALGO07")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_07_01
                    .Caption = GetText(g_arr_Text, "ALGO40")
                    .top = frmInfo.lbl_algo_07_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_07_02
                    .Caption = GetText(g_arr_Text, "ALGO41")
                    .top = frmInfo.lbl_algo_07_00.Height + frmInfo.lbl_algo_07_01.Height + 24
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .multiPage_algo(6)
                    .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                    .ScrollHeight = .lbl_algo_07_00.Height + .lbl_algo_07_01.Height + .lbl_algo_07_02.Height + 30 'height of the vertical scrolling
                    .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
                End With
            'Algorithm 08 '
                With .lbl_algo_08_00
                    .Caption = GetText(g_arr_Text, "PAGEALGO08")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_08_01
                    .Caption = GetText(g_arr_Text, "ALGO42")
                    .top = frmInfo.lbl_algo_08_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_08_02
                    .Caption = GetText(g_arr_Text, "ALGO43")
                    .top = frmInfo.lbl_algo_08_00.Height + frmInfo.lbl_algo_08_01.Height + 24
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .multiPage_algo(7)
                    .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                    .ScrollHeight = .lbl_algo_08_00.Height + .lbl_algo_08_01.Height + .lbl_algo_08_02.Height + 30 'height of the vertical scrolling
                    .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
                End With
            'Algorithm 09 '
                With .lbl_algo_09_00
                    .Caption = GetText(g_arr_Text, "PAGEALGO09")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_09_01
                    .Caption = GetText(g_arr_Text, "ALGO44")
                    .top = frmInfo.lbl_algo_09_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_09_02
                    .Caption = GetText(g_arr_Text, "ALGO45")
                    .top = frmInfo.lbl_algo_09_00.Height + frmInfo.lbl_algo_09_01.Height + 24
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .multiPage_algo(8)
                    .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                    .ScrollHeight = .lbl_algo_09_00.Height + .lbl_algo_09_01.Height + .lbl_algo_09_02.Height + 30 'height of the vertical scrolling
                    .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
                End With
                
        'Page "Code"
            With .lbl_info_code01
                .Caption = GetText(g_arr_Text, "CODE01")
                .Font.Size = 12
                .WordWrap = True
                .AutoSize = True 'label height depending on the content
            End With
            .btn_info_code01.ControlTipText = GetText(g_arr_Text, "CODE02")
            With .lbl_info_code03
                .Caption = GetText(g_arr_Text, "CODE03")
                .Font.Size = 12
                .WordWrap = True
                .AutoSize = True 'label height depending on the content
            End With
            With .lbl_info_code04
                .Caption = GetText(g_arr_Text, "CODE04") & vbNewLine & vbNewLine _
                            & GetText(g_arr_Text, "CODE05") & vbNewLine & vbNewLine _
                            & GetText(g_arr_Text, "CODE06") & vbNewLine & vbNewLine _
                            & GetText(g_arr_Text, "CODE07")
                .WordWrap = True
                .AutoSize = True 'label height depending on the content
            End With
            With .multiPage_info.Pages(3)
                .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                .ScrollHeight = .lbl_info_code04.Height 'height of the vertical scrolling
                .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
            End With

        'Page "Contact"
            With .lbl_info_contact01a
                .Caption = GetText(g_arr_Text, "CON01") & vbNewLine _
                            & GetText(g_arr_Text, "CON02")
                .WordWrap = True
            End With
            .lbl_info_contact01b.Caption = GetText(g_arr_Text, "CON03a")
            .lbl_info_contact01c.Caption = GetText(g_arr_Text, "CON03b")
            With .btn_info_contact01
                .Caption = GetText(g_arr_Text, "CON04")
                .ControlTipText = GetText(g_arr_Text, "CON05")
                .WordWrap = True
            End With
            .btn_info_contact02.ControlTipText = GetText(g_arr_Text, "CON06")
            With .lbl_info_contact02
                .Font.Size = 12
                .TextAlign = fmTextAlignRight
                .Caption = GetText(g_arr_Text, "CON07")
            End With
            With .btn_info_contact03
                .Caption = GetText(g_arr_Text, "CON08")
                .ControlTipText = GetText(g_arr_Text, "CON09")
                .WordWrap = True
            End With
            With .btn_info_contact04
                .Caption = GetText(g_arr_Text, "CON10")
                .ControlTipText = GetText(g_arr_Text, "CON11")
                .WordWrap = True
            End With
            With .btn_info_contact05
                .Caption = GetText(g_arr_Text, "CON12")
                .ControlTipText = GetText(g_arr_Text, "CON13")
                .WordWrap = True
            End With
            With .lbl_info_contact03
                .ControlTipText = GetText(g_arr_Text, "CON14")
                .WordWrap = True
            End With

        'Page "Donation"
            With .lbl_info_donation01
                .Font.Size = 12
                .Caption = GetText(g_arr_Text, "DON01") & vbNewLine & vbNewLine _
                            & GetText(g_arr_Text, "DON02")
                .AutoSize = True 'label height depending on the content
            End With
            .btn_info_donation01.ControlTipText = GetText(g_arr_Text, "DON03")
            With .btn_info_donation02
                .Caption = GetText(g_arr_Text, "DON04")
                .Font.Size = 24
                .ControlTipText = GetText(g_arr_Text, "DON05")
            End With
        
        'Page "Privacy Policy"
            With .lbl_info_privacy01
                .Caption = GetText(g_arr_Text, "PRIVACY01") & " " _
                            & GetText(g_arr_Text, "PRIVACY02")
                .WordWrap = True
            End With
            
        .show (vbModal) 'modal
    End With
End Sub

'Warnhinweis-Button im Menüband
Private Sub AI_Warning(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call warning
End Sub

Private Sub warning()
    Call AssignDataSheets
    Call GetTextComponents
    Call ShowWarning
End Sub

'Movie button (Ribbon)
Private Sub AI_Movie2017(control As IRibbonControl)

    'Check whether algorithms are allowed
    If objOption.STOP_ALG Then
        If basAuxiliary.AllowAlgorithms = False Then
            objOption.STOP_ALG = False
        Else
            Exit Sub
        End If
    End If
    
    Call AssignDataSheets
    Call GaloppSimMovie2017
End Sub

Private Sub GaloppSimMovie2017()
    'Close pop-ups if visible
        If frmBettingAnalysis.Visible Then Unload frmBettingAnalysis 'betting analysis
        If frmRS_navigation.Visible Then Unload frmRS_navigation 'navigation panel (RS edition only)
    'Play the movie
    Call basMovie2017.PlayMovie2017
End Sub

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

Public Sub TitleScreen()

    'Leave the venue?
    If objRace.STARTED Then
        'Set the button mode
        g_strMsgButtons = "CancelOK"
        'Assign the text for the pop-up
        g_strMsgCaption = g_c_tool
        g_strMsgText = GetText(g_arr_Text, "WARN003")
        'Display the pop-up
        frmMsg_MultiPurpose.show (vbModal) 'modal
        'Evaluate the return value
        If g_strButtonPressed = "CANCEL" Then Exit Sub
    End If

    Unload frmBettingAnalysis
    If objRace.STARTED Then objRace.STARTED = False
    
    'Reset Excel ribbon
    g_RibbonGaloppSim.Invalidate

    Call AssignDataSheets
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

'Ende-Button im Menüband
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
    Unload frmBettingAnalysis 'ToDo sauber machen, evtl. in eigener prozedur
    
    'Reset Excel ribbon
    g_RibbonGaloppSim.Invalidate

    'Reset Excel options
    Call ResetExcelOptions
    Application.ScreenUpdating = True
End Sub

'Language button "DE"
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

'Language button "EN"
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

''Language button "RU"
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
''Language button "CH"
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
'Language button "BG"
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

Private Sub ChangeLanguage()
    Dim oleObj As OLEObject
    
    Call GetTextComponents 'Get new texts
    Call AnimalGrammar
    
    If g_strPlayMode = "RS" Then
        For Each oleObj In g_wksRace.OLEObjects
            If oleObj.name <> "CBraces" Then 'no action at dropdown with the races
                Call RS_RefreshButtonTexts(oleObj.name)
            End If
        Next oleObj
    Else 'AI mode
        g_RibbonGaloppSim.Invalidate
    End If

End Sub

Private Sub RS_RefreshButtonTexts(name As String)

    Dim captionStart As String
    captionStart = basAuxiliary.getCaptionStartBtn(objOption.BET_MODE)
    
    Select Case name
        'Button labels
        Case "startrace" 'ToDO.. With g_wksRace.OLEObjects(name).Object.Caption .... End With
            g_wksRace.OLEObjects(name).Object.Caption = GetText(g_arr_Text, captionStart)
        Case "finishphoto"
            g_wksRace.OLEObjects(name).Object.Caption = GetText(g_arr_Text, "BTN004")
        Case "results"
            g_wksRace.OLEObjects(name).Object.Caption = GetText(g_arr_Text, "BTN005")
        Case "winner"
            g_wksRace.OLEObjects(name).Object.Caption = GetText(g_arr_Text, "BTN006")
        Case "bets"
            g_wksRace.OLEObjects(name).Object.Caption = GetText(g_arr_Text, "BTN007")
        Case "raceoptions"
            g_wksRace.OLEObjects(name).Object.Caption = GetText(g_arr_Text, "BTN001")
        Case "exceloptions"
            g_wksRace.OLEObjects(name).Object.Caption = GetText(g_arr_Text, "BTN002")
        Case "language"
            g_wksRace.OLEObjects(name).Object.Caption = GetText(g_arr_Text, "LANGUAGE001")
        Case "info"
            g_wksRace.OLEObjects(name).Object.Caption = GetText(g_arr_Text, "BTN009")
        Case "warning"
            g_wksRace.OLEObjects(name).Object.Caption = GetText(g_arr_Text, "BTN010")
        Case "movie2017"
            g_wksRace.OLEObjects(name).Object.Caption = GetText(g_arr_Text, "BTN011")
    End Select
End Sub

'Anzahl der installierten Rennen auslesen
Private Sub AI_InstalledRaces_getItemCount(control As IRibbonControl, ByRef returnedVal)
    
    Dim wksR As Worksheet
    Dim cnt As Long
    
    Set m_colRacesInstalled = Nothing
    Set m_colRacesInstalled = New Collection
    
    For Each wksR In ThisWorkbook.Worksheets
        If left(wksR.name, 5) = "race_" Then cnt = cnt + 1
    Next wksR
    
    returnedVal = cnt
    
End Sub

'Namen der installierten Rennen auslesen
Private Sub AI_InstalledRaces_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)

    Dim wksCheck As Worksheet
    Dim cnt As Long
    
    For Each wksCheck In ThisWorkbook.Worksheets
        If left(wksCheck.name, 5) = "race_" Then cnt = cnt + 1
        If cnt = index + 1 Then
            With wksCheck
                m_colRacesInstalled.Add .name
                returnedVal = .Cells(basAuxiliary.GetRow(wksCheck, "RACE NAME"), 2).Value & " " & _
                                .Cells(basAuxiliary.GetRow(wksCheck, "YEAR"), 2).Value & " (" & _
                                .Cells(basAuxiliary.GetRow(wksCheck, "DISTANCE METERS"), 2).Value & "m) - " & .Cells(basAuxiliary.GetRow(wksCheck, "TRACK LOCATION"), 2).Value
                End With
            Exit For
        End If
    Next wksCheck

End Sub

'Set default race
Private Sub AI_InstalledRaces_GetSelectedItemID(control As IRibbonControl, ByRef itemID As Variant)
    If objRace.SELECTED = "" Then
        objRace.SELECTED = m_colRacesInstalled(1) 'take the first race of the collection
    End If
    itemID = objRace.SELECTED
End Sub

'Ausgewähltes Rennen setzen
Private Sub AI_InstalledRaces_Click(control As IRibbonControl, id As String, index As Integer)
    objRace.SELECTED = m_colRacesInstalled(index + 1)
End Sub

'Execute this procedure when the workbook is being closed.
'The Auto_Close procedure can be used alternatively to the Workbook_BeforeClose
'event in "ThisWorkbook" ("DieseArbeitsmappe") which is NOT used for this project.
'If both procedures are implemented first the Workbook_BeforeClose is executed
'followed by Auto_Close.
Public Sub Auto_Close()
    'Reset Excel options
        Call basMainCode.ResetExcelOptions
        Application.ScreenUpdating = True
    If g_strPlayMode = "RS" Then
        'Do not save the workbook
        'https://support.microsoft.com/de-de/help/213428/how-to-suppress-save-changes-prompt-when-you-close-a-workbook-in-excel
        ThisWorkbook.Saved = True
    End If
End Sub
