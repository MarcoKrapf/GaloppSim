Attribute VB_Name = "basMainCode"
Option Explicit
Option Private Module

'This module contains the main code of the race simulator with most of the logic

'GALOPPSIM - Version 149.00 - September 2018
'Horse racing simulator for Microsoft Excel
'Author: Marco Matjes - info@galoppsim.racing - https://galoppsim.racing/
'License: GNU General Public License v3.0

'Naming conventions
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

'Global constants and variables
Public Const g_c_tool As String = "GaloppSim" 'Name of the tool
Public Const g_c_version As String = "(v149.00)" 'Version of the tool
Public Const g_c_email As String = "info@galoppsim.racing" 'Contact e-mail address
Public Const g_c_defaultRaceOptionsFile As String = "RaceOptions" 'File name for the race options
Public Const g_c_defaultFileType As String = ".galoppsim" 'File type for GaloppSim files
Public g_defaultPath As String 'Path for the Galoppsim files
Public g_RibbonGaloppSim As IRibbonUI 'Custom ribbon
Public g_strLanguage As String 'Interface Language
Public g_blnSpeech As Boolean 'Speech or not
Public g_strPlayMode As String ' "AI" = AddIn (.xlam) / "RS" = Run Simple (.xlsm)
Public g_objLabel As MSForms.Label 'Label used for different purposes
Public g_colRSbuttons As Collection 'Menu buttons in the RS edition
Public g_oleComboRaces As OLEObject 'ComboBox with installed races in the RS edition
Public g_blnStopAlgorithms As Boolean 'If true: Algorithms at work!
Public g_strRaceID As String 'Unique race ID
Public g_blnRaceStarted As Boolean 'Flag indicating whether a race has been started
Public g_byteTrackZoom As Byte 'Track zoom level (1=small 2=medium 3=large)
Public g_intTrackMetres As Integer 'Meteranzeige an der Rennbahn
Public g_blnHorseNamesLeft As Boolean 'Flag indicating whether the names of the horses are permanently displayed at the left margin
Public g_blnHorseColoursLeft As Boolean 'Flag indicating whether the colours of the horse are displayed on the left margin
Public g_blnHorseNamesFinish As Boolean 'Flag indicating whether the names of the horses are displayed in the destination
Public g_blnHoofprints As Boolean 'Flag indicating whether hoof prints are displayed
Public g_blnTactics0 As Boolean 'If true: No racing tactics
Public g_blnTactics3 As Boolean 'If true: Racing tactics (three phases)
Public g_blnTactics6 As Boolean 'If true: Racing tactics (six phases)
Public g_blnIncidentRefuse As Boolean 'Flag indicating whether horses can refuse to run
Public g_blnSlipstream As Boolean 'Flag indicating whether the slipstream effect is active
Public g_blnSlipstreamDouble As Boolean 'Flag indicating whether the double slipstream effect is selected
Public g_blnSlipstreamShow As Boolean 'Flag indicating whether the slipstream effect is displayed graphically
Public g_blnFocusedRun As Boolean 'Flag indicating whether the Focused Run mode is active
Public g_blnHighlightFoc As Boolean 'Flag indicating whether the horse in focus is highlighted during the race
Public g_intFocusedRun As Integer 'Number of the focused horse
Public g_blnRankingDelay As Boolean 'Flag indicating whether the results are displayed on the ranking list bottom up with delay
Public g_blnRankingColours As Boolean 'Flag indicating whether the horse colours are displayed on the ranking list
Public g_blnHighlightFav As Boolean 'Flag indicating whether the favorite is highlighted during the race
Public g_blnRaceInformation As Boolean 'Flag indicating whether race information is displayed during the race
Public g_blnRaceInfoPopup As Boolean 'Flag indicating whether the race information is displayed in a pop-up
Public g_blnRaceInfoWorksheet As Boolean 'Flag indicating whether the race information is displayed on the worksheet
Public g_blnRaceInfoLeader As Boolean 'Flag indicating whether the name of the current leader is displayed
Public g_blnRaceInfoProgressBar As Boolean 'Flag indicating whether a progress bar of the race distance is displayed
Public g_lngRaceInfoBackColour As Long 'Background colour of the race information
Public g_lngRaceInfoForeColour As Long 'Foreground colour of the race information
Public g_blnBettingMode As Boolean 'Flag indicating whether bettings can be placed
Public g_blnBettingAnalysis As Boolean 'Flag indicating whether a betting analysis is performed automatically after the race
Public g_blnBetsPlaced As Boolean 'If true: Bets have been placed
Public g_intGMPspeedfactor As Integer 'Factor for the race speed
Public g_arr_RaceSpeed(0 To 4) As Integer 'array with the race speed factors
Public g_strRaceSelected As String 'Selected race
Public g_arr_varHorses() As Variant 'All information about the horses
Public g_colBetSlips As Collection 'List of all betting slips
Public g_intHorsesEnrolled As Integer 'number of horses registered (including horses that don´t start)
Public g_arrTxt() As Variant 'Text components (general)
Public g_arrTxtInfo() As Variant 'Text components (info section)
Public g_strMsgCaption As String 'Caption for the UserForm
Public g_strMsgText As String 'Text for the UserForm
Public g_strMsgButtons As String 'Buttons for the UserForm
Public g_strPlayerName As String 'Name of the player who places a bet
Public g_strButtonPressed As String 'Return value of the pressed button
Public g_strExcelMode As String 'AI edition: Excel options - Excel mode, TV mode with menu strip or TV mode full-screen

Public g_wksRace As Worksheet 'Worksheet for the race
Public g_wksMovie As Worksheet 'Worksheet for the movie
Public g_wksRaceData As Worksheet 'Worksheet with the race data
Public g_wksPicData As Worksheet 'Worksheet with the picture data

'Variables on module-level
Dim m_wksTxt As Worksheet 'Worksheet with the text components
Dim m_wksTxtCountry As Worksheet 'Worksheet with the country names
Dim m_wksTxtInfo As Worksheet 'Worksheet with the info text
Dim m_wksTec As Worksheet 'Worksheet with technical data (speed, tactics...)
Dim m_wksAdv As Worksheet 'Worksheet with advertising data (.gsadv file format)
Dim m_wksCheck As Worksheet 'Variable to check whether the table sheet GALOPPSIM exists
Dim m_arr_varPhotofinish() As Variant 'Position of each horse for the finish photo
Dim m_arr_varResultsCalc() As Variant 'Calculation of the position at the finish line
Dim m_arr_varResults() As Variant 'Result list
Dim m_colRacesInstalled As Collection 'List of installed races
Dim m_arr_varAdv() As Variant 'Advertising sequence
Dim m_strRaceName As String 'name of the race
Dim m_strRealRace As String 'race that took place in reality (Y/N)
Dim m_strRaceLocation As String 'location of the race track
Dim m_strRaceCountry As String 'country of the race track
Dim m_strRaceTrack As String 'name of the race track
Dim m_strRaceYear As String 'year of the race
Dim m_strTrackSurface As String 'type of the race track (T = turf, D = dirt, S = snow, W = water...)
Dim m_lngTrackColour As Long 'colour of the track surface
Dim m_strRaceType As String 'race type (F = flat, S = steeplechase...)
Dim m_strRandomLane As String 'lanes fix or random (F/R)
Dim m_strRandomColour As String 'horse colours fix or random (F/R)
Dim m_strRandomOdd As String 'odds fix or random (F/R)
Dim m_strAdvertising As String 'advertising (Y/N)
Dim m_intTopRows As Integer 'number of rows at the top of the worksheet (used for the menu in the RS edition)
Dim m_intLeftColumns As Integer 'number of columns left of the start boxes
Dim m_intColumsAfterFinish As Integer 'number of columns behind the finish line
Dim m_intTrackLength As Integer 'length of the track in meters
Dim m_intTrackCellHeight As Integer 'cell height (race track)
Dim m_dblTrackCellWidth As Double 'cell width (race track)
Dim m_dblRankingsWidth As Double 'cell width of the finish photo and the ranking list
Dim m_intAdvertisingHeight As Integer 'row height of the advertising area
Dim m_intFontSize As Integer 'font size of the horse names and the hoof prints
Dim m_lngSpeedCondHigh As Long, m_lngSpeedCondLow As Long 'variables for the speed range of the daily form of the horses
Dim m_lngSpeedLoopHigh As Long, m_lngSpeedLoopLow As Long 'variables for the range of the randomly assigned speed per step
Dim m_lngSpeedTacticsHigh As Long, m_lngSpeedTacticsMedium As Long, m_lngSpeedTacticsLow As Long 'variables for the speed range of each phase if racing tactics are active
Dim m_intHorsesStarting As Integer 'number of horses that start
Dim m_intHorsesRunning As Integer 'number of horses currently running
Dim m_byteFavourite(1 To 3) As Byte 'array for three predicted favourites of the race
Dim m_dblFavCalc(1 To 3) As Double 'array for calculating the favourites
Dim m_intHorsesFinishing As Integer 'variable that counts how many horses arrive at the finish line in one loop
Dim m_intFinishLoop As Integer 'variable that counts the finishing loop in which a placement was calculated
Dim m_intPlace As Integer 'placement in the finish
Dim m_strWinner As String 'name(s) of the horse(s) in 1st place
Dim m_blnDeadHeat As Boolean 'dead heat (more than one horse has won)
Dim m_blnWin As Boolean 'flag indicating whether a horse has won the race
Dim m_blnPhotofinish As Boolean 'flag indicating whether there is a photofinish
Dim m_blnExactFinish As Boolean 'variable for the exact calculation at the finish line
Dim m_intPosLeader As Integer 'position of the leading horse
Dim m_strNameLeader As String 'name of the leading horse
Dim i As Integer, j As Integer, k As Integer, m As Integer 'counting variables for loops
Dim z As Long 'auxiliary variable for loops

'Starting procedure (RS edition)
Public Sub NewRace_RS()
    
    ThisWorkbook.Worksheets("RS").Activate
    m_intTopRows = 8
    
    Call AssignDataSheets 'Assign worksheets "Txt", "Tec", "Adv" and "Pic"
    Call GetTextComponents 'Get texts
    Call CreateRaceSheet 'Create worksheet "GALOPPSIM"
    Call RS_StartScreen 'Startscreen with Loppsi and navigation panel
    Call RS_AddCommandButtons 'Add controls
    Call RS_AddComboBox 'Add dropdown for race selection
End Sub

'Starting procedure
Private Sub NewRace()

    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
        'Close pop-ups if visible
        If frmBettingAnalysis.Visible Then Unload frmBettingAnalysis 'betting analysis
        If frmRS_navigation.Visible Then Unload frmRS_navigation 'navigation panel (RS edition only)
        
        If g_strPlayMode = "AI" Then
            Call AssignDataSheets 'Assign worksheets with texts, pictures and technical data
            Call GetTextComponents 'Texte einlesen
        End If
        
        'Reset the betting slip collection
        Set g_colBetSlips = Nothing
        Set g_colBetSlips = New Collection
    
        Call GetRaceData 'Tabellenblatt mit ausgewähltem Rennen
        Call AssignBasicValues 'Grundsätzliche Daten auslesen bzw. festlegen
        Call GetHorseData 'Daten über das Rennen einlesen
        Call UserFormSTART 'Show Start-UserForm
        
    If g_blnRaceStarted Then
        If g_strPlayMode = "AI" Then
            Call CreateRaceSheet 'Tabellenblatt "GALOPPSIM"
            Call basAuxiliary.AI_ExcelModeStart
            Call basAuxiliary.AI_ExcelModeEnd
            Call CursorAway 'Place the cursor far away (in the upper right corner of the screen)
        End If
        
        If g_strPlayMode = "RS" Then
            Cells.Clear 'Clear the whole worksheet
            With g_wksRace.Cells(2, 2) 'Write title
                .Font.name = "Arial Black"
                .Value = g_c_tool & " " & g_c_version
            End With
            Call RS_HideNavi 'hide the navigation area
        End If

        Call DrawRaceTrack 'Geläuf zeichnen
        Call DrawHorseNames 'Pferdenamen am Start und im Ziel wenn angehakt
        If g_blnRaceInformation And g_blnRaceInfoPopup Then Call RaceInfoPopup 'Show pop-up with race info if checked
        Call RaceWelcome 'Popup zu Rennbeginn
        Call StartingGrid 'Pferde in Boxen stellen
        Call RacePresentation 'Pferde vorstellen
        Call RunRace 'Rennstart
        Call NotFinished 'Find the horses that did not finish
        Call ShowRankings 'Ergebnistafel
        Call DrawWinner 'Grafik
        If g_blnBetsPlaced And g_blnBettingAnalysis Then Call UserFormAnalyseBetSlips 'Analyse bet slips
        
        If g_strPlayMode = "RS" Then 'Show the navigation area when running RS mode
            'Activate buttons
            With g_wksRace
                .OLEObjects("fotofinish").Object.Enabled = True
                .OLEObjects("results").Object.Enabled = True
                .OLEObjects("winner").Object.Enabled = True
                If g_blnBetsPlaced Then
                    .OLEObjects("bets").Object.Enabled = True
                End If
            End With
            Call RS_ShowNavigationPanel(True)
        End If

        If g_strPlayMode = "AI" Then
            g_RibbonGaloppSim.Invalidate 'refresh the status buttons
            Call basAuxiliary.AI_ExcelModeEnd
        End If
        
    End If

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

Public Sub RS_AddComboBox()
    'Add combobox to the navigation panel
    Call RS_addComboboxRaces("CBraces", 196, 15, 419, 22) '"name(ID)", left, top, width, height
End Sub

'Add controls in RS mode
Public Sub RS_AddCommandButtons()
    'Error handling
    On Error GoTo ERRORHANDLING

    Dim captionStart As String
    captionStart = basAuxiliary.getCaptionStartBtn(g_blnBettingMode)
    
    'Add buttons to the navigation panel
        Set g_colRSbuttons = New Collection
        
        '"name(ID)", left, top, width, height, font-size, font:bold, _
            background-color (hex), text(xxx is the initial caption)
        Call RS_addButton("raceoptions", 15, 40, 81, 49, 11, False, &HFFFFFF, GetTxt(g_arrTxt, "BTN001"))
        Call RS_addButton("exceloptions", 99, 40, 81, 49, 11, False, &HFFFFFF, GetTxt(g_arrTxt, "BTN002"))
        Call RS_addButton("startrace", 196, 40, 81, 49, 14, True, &HC000&, GetTxt(g_arrTxt, captionStart))
        Call RS_addButton("fotofinish", 280, 40, 81, 49, 11, False, &HFFFFFF, GetTxt(g_arrTxt, "BTN004"))
        Call RS_addButton("results", 364, 40, 81, 49, 11, False, &HFFFFFF, GetTxt(g_arrTxt, "BTN005"))
        Call RS_addButton("winner", 448, 40, 81, 49, 11, False, &HFFFFFF, GetTxt(g_arrTxt, "BTN006"))
        Call RS_addButton("bets", 532, 40, 81, 49, 11, False, &HFFFFFF, GetTxt(g_arrTxt, "BTN007"))
        Call RS_addButton("language", 629, 40, 81, 49, 11, False, &HFFFFFF, GetTxt(g_arrTxt, "LANGUAGE001"))
        Call RS_addButton("info", 713, 40, 81, 49, 11, False, &HFFFFFF, GetTxt(g_arrTxt, "BTN009"))
        Call RS_addButton("warning", 797, 40, 81, 49, 11, False, &HFFFFFF, GetTxt(g_arrTxt, "BTN010"))
        Call RS_addButton("movie2017", 881, 40, 81, 49, 11, False, &HFFFFFF, GetTxt(g_arrTxt, "BTN011"))

        Call RS_InactivateCommandButtons 'inactivate some buttons
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

Private Sub RS_InactivateCommandButtons()
    'Set buttons inactive
        With g_wksRace
            .OLEObjects("fotofinish").Object.Enabled = False
            .OLEObjects("results").Object.Enabled = False
            .OLEObjects("winner").Object.Enabled = False
            .OLEObjects("bets").Object.Enabled = False
        End With
End Sub

'Click on a button in the Run Simple Edition
Public Sub RS_execute_Click(name As String)
    
    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    Select Case name

        Case "startrace"
            If g_oleComboRaces.Object.Value = "" Then

            Else
                'Leave the current race?
                If g_blnRaceStarted Then
                    
'                    'A MessageBox cannot handle unicode, so for example Cyrillic characters are displayed as question marks
'                    If MsgBox((GetTxt(g_arrTxt, "RACE003") & ": " & g_wksRace.OLEObjects("CBraces").Object.text), _
'                        vbOKCancel, g_c_TOOL) = vbCancel Then Exit Sub
                        
                    'Set the button mode
                    g_strMsgButtons = "CancelOK"
                    'Assign the text for the pop-up
                    g_strMsgCaption = g_c_tool
                    g_strMsgText = GetTxt(g_arrTxt, "RACE003") & ": " & g_wksRace.OLEObjects("CBraces").Object.text
                    'Display the pop-up
                    frmMsg_MultiPurpose.Show (vbModal) 'modal
                    'Evaluate the return value
                    If g_strButtonPressed = "CANCEL" Then Exit Sub
                    
                    Call ShowNewRaceScreen
                End If
                g_strRaceSelected = g_oleComboRaces.Object.Value 'get m_strRaceName from dropdown
            End If
            
            Call RS_InactivateCommandButtons 'inactivate some buttons
            Call NewRace
        
        Case "fotofinish"
            Call show_foto
        Case "results"
            Call show_ergebnis
        Case "winner"
            Call show_winnerPhoto
        Case "bets"
            Call show_wetten
        Case "raceoptions"
            frmOptionsRace.Show (vbModal) 'display UserForm (modal)
        Case "exceloptions"
            frmOptionsExcel.Show (vbModal) 'display UserForm (modal)
        Case "language"
            frmRS_languages.Show (vbModal) 'display UserForm (modal)
            Call ChangeLanguage
        Case "info"
            Call show_info
        Case "warning"
            Call show_warning
        Case "movie2017"
            Call GaloppSimMovie2017
    End Select
    
    Exit Sub

End Sub

Private Sub ShowNewRaceScreen()
    With g_wksRace.UsedRange
        .ColumnWidth = ZoomLevelPictures()(0)
        .RowHeight = ZoomLevelPictures()(1)
        .Clear
    End With
    If g_strPlayMode = "RS" Then Call RS_HideNavi 'hide the navigation area
    If g_strPlayMode = "AI" Then Call basAuxiliary.AI_ExcelModeStart
    ActiveWindow.ScrollColumn = 1 'scroll to the left
    Call PaintPicture("NEWRACE", 100, 40, 1, 1) 'paint title picture
    'Place the cursor far away (in the upper right corner of the screen)
        Call CursorAway
End Sub

Private Sub PaintPicture(pic As String, cols As Integer, rows As Integer, top As Integer, left As Integer)

    'Deactivate screen updating
        Application.ScreenUpdating = False
        
    'Paint a picture (get data from the worksheet "Pic")
        k = basAuxiliary.GetPictureColumn(pic)
        
        m = 2 'initial row for reading the picture
        
        For i = top To rows
            For j = left To cols
                g_wksRace.Cells(i, j).Interior.Color _
                    = g_wksPicData.Cells(m, k).Value
                m = m + 1 'next row on the worksheet "Pic"
            Next j
        Next i

    'Activate screen updating
        Application.ScreenUpdating = True
        
End Sub

'Add new RS button
Private Sub RS_addButton(n As String, l As Integer, t As Integer, w As Integer, _
                            h As Integer, fs As Integer, fb As Boolean, bc As Long, c As String)
    Dim oleRSbutton As OLEObject
    Dim objRSbutton As clsRSbutton
    
    Set oleRSbutton = g_wksRace.OLEObjects.Add(classtype:="Forms.CommandButton.1", _
    left:=l, top:=t, Width:=w, Height:=h)
    
    With oleRSbutton
        .name = n 'ID
        .Object.Font.Size = fs
        .Object.Font.Bold = fb
        .Object.Caption = c
        .Object.BackColor = bc
        .Object.WordWrap = True
        .Object.TakeFocusOnClick = False
        .Placement = xlFreeFloating
        .Visible = True
    End With
    
    Set objRSbutton = New clsRSbutton
    Set objRSbutton.RSButtonObject = oleRSbutton.Object
    objRSbutton.RSbtnID = n

    g_colRSbuttons.Add objRSbutton
 
End Sub

'Add new RS combobox with races
Private Sub RS_addComboboxRaces(n As String, l As Integer, t As Integer, w As Integer, h As Integer)

    Dim wksCheck As Worksheet
    
    Set m_colRacesInstalled = Nothing
    Set m_colRacesInstalled = New Collection

    Set g_oleComboRaces = g_wksRace.OLEObjects.Add(classtype:="Forms.ComboBox.1", _
        left:=l, top:=t, Width:=w, Height:=h)
    
    With g_oleComboRaces
        .name = n 'ID
        .Placement = xlFreeFloating
        .Object.ColumnCount = 2 'column 0: g_wksRace-name // column 1: visible name
        .Object.ColumnWidths = "0 Pt" 'width of the column with the m_strRaceName name --> hidden
        .Object.Style = fmStyleDropDownList 'Allow only values from the item list, no free entries
        .Visible = True
    End With
    
    'Populate Dropdown with installed races
    For Each wksCheck In ThisWorkbook.Worksheets
        If left(wksCheck.name, 5) = "race_" Then
            m_colRacesInstalled.Add wksCheck.name
            With g_oleComboRaces.Object
            .AddItem
            .List(.ListCount - 1, 0) = wksCheck.name
            .List(.ListCount - 1, 1) = wksCheck.Cells(4, 2).Value & " " & _
                                    wksCheck.Cells(5, 2).Value & " (" & _
                                    wksCheck.Cells(12, 2).Value & "m) - " & wksCheck.Cells(6, 2).Value
            End With
        End If
    Next wksCheck
    
    'Set default race
    g_oleComboRaces.Object.ListIndex = 0 'take the first race

End Sub

Private Sub RS_HideNavi()
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

    If popup = True Then frmRS_navigation.Show (vbModeless) 'modeless

End Sub

'Start screen for the RS edition
Private Sub RS_StartScreen()
        With g_wksRace.Range(Columns(1), Columns(100)) 'set the column width and row height dependent on the screen size
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
    'Paint title picture
        Call PaintPicture("RUNSIMPLE", 100, 40, 1, 1)
    'Place the cursor far away (in the upper right corner of the screen)
        Call CursorAway
End Sub

'Create worksheet for the race
Private Sub CreateRaceSheet()
    'In case an error occurs
    On Error GoTo ERRORHANDLING

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
    On Error GoTo ERRORHANDLING
    
    With ThisWorkbook
        Set m_wksTxt = .Worksheets("Txt")
        Set m_wksTxtCountry = .Worksheets("TxtCountries")
        Set m_wksTxtInfo = .Worksheets("TxtInfo")
        Set m_wksTec = .Worksheets("Tec")
        Set m_wksAdv = .Worksheets("Adv")
        Set g_wksPicData = .Worksheets("Pic")
    End With
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Grunddaten einlesen
Private Sub AssignBasicValues()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
        
    'Determine the zoom level
    If g_byteTrackZoom = 0 Then
        g_byteTrackZoom = ZoomLevelRecommendation()
    Else
        Dim byteOpt As Byte
        byteOpt = ZoomLevelRecommendation() 'get the perfect level by calling the function

        If g_byteTrackZoom <> byteOpt Then 'compare the selected value with the recommendation
                    'Set the button mode
                    g_strMsgButtons = "YesNo"
                    'Assign the text for the pop-up
                    g_strMsgCaption = m_strRaceName & " " & m_strRaceYear
                    g_strMsgText = GetTxt(g_arrTxt, "ZOOM007") & vbNewLine & vbNewLine & _
                                    GetTxt(g_arrTxt, "ZOOM008") & ": " & ZoomLevelText(g_byteTrackZoom) & vbNewLine & _
                                    GetTxt(g_arrTxt, "ZOOM009") & ": " & ZoomLevelText(byteOpt) & vbNewLine & vbNewLine & _
                                    GetTxt(g_arrTxt, "ZOOM010")
                    'Display the pop-up
                    frmMsg_MultiPurpose.Show (vbModal) 'modal
                    'Evaluate the return value
                    If g_strButtonPressed = "YES" Then g_byteTrackZoom = byteOpt 'adapt the value
        End If
    End If
    
    'Assign the values for the zoom level
        Select Case g_byteTrackZoom
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

'Grunddaten über das Renneneinlesen
Private Sub GetRaceData()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    'Tabellenblatt mit ausgewähltem Rennen zuweisen
    Set g_wksRaceData = ThisWorkbook.Worksheets(g_strRaceSelected)
    'Grunddaten aus Tabellenblatt einlesen
    With g_wksRaceData
        g_strRaceID = .Cells(1, 2).Value 'unique race ID
        m_strRealRace = .Cells(2, 2).Value 'real race (yes or no)
        m_strRaceName = .Cells(4, 2).Value 'race name
        m_strRaceYear = .Cells(5, 2).Value 'year of the race
        m_strRaceLocation = .Cells(6, 2).Value 'track location
        m_strRaceCountry = GetCountryName(.Cells(7, 2), g_strLanguage) 'country
        m_strRaceTrack = .Cells(8, 2).Value 'track name
        m_lngTrackColour = .Cells(9, 2).Value 'track colour
        Select Case .Cells(10, 2).Value 'track surface
            Case "T" 'turf
                m_strTrackSurface = GetTxt(g_arrTxt, "TRACK001")
            Case "D" 'dirt
                m_strTrackSurface = GetTxt(g_arrTxt, "TRACK002")
            Case "S" 'snow
                m_strTrackSurface = GetTxt(g_arrTxt, "TRACK003")
        End Select
        Select Case .Cells(11, 2).Value 'race type
            Case "F" 'flat race
                m_strRaceType = GetTxt(g_arrTxt, "RACETYPE01")
            Case "S" 'steeplechase
                m_strRaceType = GetTxt(g_arrTxt, "RACETYPE02")
            Case Else
        End Select
        m_intTrackLength = .Cells(12, 2).Value 'race distance
        g_intHorsesEnrolled = .Cells(14, 2).Value 'number of horses
        m_strRandomLane = .Cells(15, 2).Value 'lanes fix or random
        m_strRandomColour = .Cells(16, 2).Value 'horse colours fix or random
        m_strRandomOdd = .Cells(17, 2).Value 'odds fix or random
        m_strAdvertising = .Cells(18, 2).Value 'advertising (yes or no)
    End With
    
    'Advertisement data
    If m_strAdvertising = "Y" Then
        j = g_wksRaceData.Cells(rows.Count, 13).End(xlUp).row - 1 'Last row
        ReDim m_arr_varAdv(1 To j) 'Location of the advertisement data
        For i = 1 To j
            For k = 1 To m_wksAdv.Cells(1, Columns.Count).End(xlToLeft).Column
                If g_wksRaceData.Cells(i + 1, 13).Value = m_wksAdv.Cells(1, k).Value Then
                    m_arr_varAdv(i) = k 'Assign column number
                    Exit For
                End If
            Next k
        Next i
    End If
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Get the country name dependent on the selected race and language
Private Function GetCountryName(code As String, language As String) As String
    Dim col As Integer, row As Integer
    
    'Find the column on the worksheet
        For col = 2 To m_wksTxtCountry.Cells(2, Columns.Count).End(xlToLeft).Column
            If m_wksTxtCountry.Cells(1, col).Value = language Then Exit For
        Next col
        
    'Find the row on the worksheet
        For row = 2 To m_wksTxtCountry.Cells(rows.Count, 2).End(xlUp).row
            If m_wksTxtCountry.Cells(row, 1).Value = code Then Exit For
        Next row
        
    'Return the country name
        If m_wksTxtCountry.Cells(row, col).Value = "" Then
            GetCountryName = m_wksTxtCountry.Cells(row, 2).Value 'return the English name if no name was found
        Else
            GetCountryName = m_wksTxtCountry.Cells(row, col).Value 'return the name according to the selected language
        End If
End Function

'Get the text components according to the selected language
Private Sub GetTextComponents()
    Dim col As Integer
    
    'Read the text components into arrays
    
        'General text components
        ReDim g_arrTxt(0 To 1, 1 To 1000) 'resize the array
        col = GetLanguageColumn(m_wksTxt) 'get the column with the language
        For i = 1 To UBound(g_arrTxt, 2)
            g_arrTxt(0, i) = m_wksTxt.Cells(i, 1).Value 'ID
            g_arrTxt(1, i) = m_wksTxt.Cells(i, col).Value 'value
        Next i

        'Info section texts
        ReDim g_arrTxtInfo(0 To 1, 1 To 200) 'resize the array
        col = GetLanguageColumn(m_wksTxtInfo) 'get the column with the language
        For i = 1 To UBound(g_arrTxtInfo, 2)
            g_arrTxtInfo(0, i) = m_wksTxtInfo.Cells(i, 1).Value 'ID
            g_arrTxtInfo(1, i) = m_wksTxtInfo.Cells(i, col).Value 'value
        Next i
    
End Sub

'Daten über die Pferde
Private Sub GetHorseData()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    'Anzahl der Starter aus Tabellenblatt auslesen
        m_intHorsesStarting = Application.WorksheetFunction.CountIf(g_wksRaceData.Columns(6), "START")
    'Arrays anlegen
    ReDim g_arr_varHorses(1 To g_intHorsesEnrolled, 0 To 22) 'Alle Daten der Pferde
    ReDim m_arr_varPhotofinish(1 To g_intHorsesEnrolled, 0 To 3) 'Snapshot für Zielfoto/Fotofinish
    ReDim m_arr_varResults(0 To m_intHorsesStarting, 0 To 5) 'Ergebnisliste
    
    'In case of random lanes
        If m_strRandomLane = "R" Then
            Dim boxNr As Integer
            Dim inBox As Boolean
            Dim boxenNr() As Integer
            ReDim boxenNr(1 To g_intHorsesEnrolled)
            For i = 1 To g_intHorsesEnrolled
                boxenNr(i) = i
            Next i
        End If
    
    For i = 1 To g_intHorsesEnrolled
        g_arr_varHorses(i, 0) = g_wksRaceData.Cells(1 + i, 6).Value 'Status (START, CANCELLED, REFUSED...)
        g_arr_varHorses(i, 11) = g_wksRaceData.Cells(1 + i, 5).Value 'Startnummer
        g_arr_varHorses(i, 1) = g_wksRaceData.Cells(1 + i, 7).Value 'Name des Pferds
        If m_strRandomColour = "F" Then 'fix
            g_arr_varHorses(i, 2) = g_wksRaceData.Cells(1 + i, 8).Value 'Horse colour
        Else 'random
            If g_arr_varHorses(i, 1) = "Loppsi" Then
                g_arr_varHorses(i, 2) = 192 'Loppsi is always red
            Else
                Randomize 'Zufallsgenerator zurücksetzen
                Do
                    g_arr_varHorses(i, 2) = CLng(((10 - 1 + 1) * Rnd + 1) * 1000000)
                Loop Until g_arr_varHorses(i, 2) >= 0 And g_arr_varHorses(i, 2) <= 16777215 'allowed value range
            End If
        End If
        
        If m_strRandomLane = "R" Then 'lanes random
            inBox = False
            Do Until inBox = True
                Randomize
                boxNr = (Int((g_intHorsesEnrolled - 1 + 1) * Rnd + 1)) 'Zufallszahl
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
        
        g_arr_varHorses(i, 4) = m_intLeftColumns + 5 'Fixe Startposition in der Box
        g_arr_varHorses(i, 5) = g_wksRaceData.Cells(1 + i, 9).Value 'Grundgeschwindigkeit des Pferds
                                                        '(linear von 1,50010 bis 1,49988)
        'Form des Pferds durch Zufallszahl festlegen
            Randomize 'Zufallsgenerator zurücksetzen
            g_arr_varHorses(i, 6) = (Int((m_lngSpeedCondHigh - m_lngSpeedCondLow + 1) * Rnd + m_lngSpeedCondLow) + 100000) / 100000 'Zufallszahl
        'Wettquote festlegen
            If m_strRandomOdd = "F" Then 'fix
                g_arr_varHorses(i, 17) = g_wksRaceData.Cells(1 + i, 10).Value
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
    Next i
    
    'Favoriten errmitteln aus Grundgeschwindigkeit und Form
        'Clear the entire array
            Erase m_dblFavCalc
            
        'Berechnung der drei Favoriten
            For i = 1 To g_intHorsesEnrolled
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
        If g_blnTactics3 Or g_blnTactics6 Then 'Geschwindigkeit in den Rennphasen 1-3 pro Pferd festlegen
            For i = 1 To g_intHorsesEnrolled
                Randomize 'Zufallsgenerator zurücksetzen
                k = (Int((6 - 1 + 1) * Rnd + 1)) 'Zufallszahl zwischen 1 und 6
                For j = 1 To 3
                    g_arr_varHorses(i, 11 + j) = m_wksTec.Cells(1 + j, 3 + k).Value
                Next j
            Next i
        End If
        
        If g_blnTactics6 Then 'Geschwindigkeit in den Rennphasen 1-6 pro Pferd festlegen
            For i = 1 To g_intHorsesEnrolled
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
        Dim strComboBoxValue As String
        
        dblWindowHeight = Application.ActiveWindow.Height 'window height
        
        On Error Resume Next 'skip if no race is selected
            'Get the selected race
            If g_strPlayMode = "AI" Then
                strComboBoxValue = g_strRaceSelected
            Else
                strComboBoxValue = g_oleComboRaces.Object.Value
            End If
            'Get the number of horses in the selected race
            intHorses = ThisWorkbook.Worksheets(strComboBoxValue).Cells(14, 2).Value
        On Error GoTo 0
        
        Select Case dblWindowHeight
            Case Is > 1100 'large window
                    ZoomLevelRecommendation = 3
            Case Is > 799 'window height
                If intHorses <= 16 Then 'up to 16 horses
                    ZoomLevelRecommendation = 3
                ElseIf intHorses <= 23 Then '17-23 horses
                    ZoomLevelRecommendation = 2
                Else 'more than 23 horses
                    ZoomLevelRecommendation = 1
                End If
            Case Is > 699 'window height
                If intHorses <= 13 Then 'up to 13 horses
                    ZoomLevelRecommendation = 3
                ElseIf intHorses <= 18 Then '14-18 horses
                    ZoomLevelRecommendation = 2
                Else 'more than 18 horses
                    ZoomLevelRecommendation = 1
                End If
            Case Is > 599 'window height
                If intHorses <= 9 Then 'up to 9 horses
                    ZoomLevelRecommendation = 3
                ElseIf intHorses <= 13 Then '9-13 horses
                    ZoomLevelRecommendation = 2
                Else 'more than 13 horses
                    ZoomLevelRecommendation = 1
                End If
            Case Else 'small window
                If intHorses <= 7 Then 'up to 7 horses
                    ZoomLevelRecommendation = 3
                ElseIf intHorses <= 10 Then '7-10 horses
                    ZoomLevelRecommendation = 2
                Else 'more than 13 horses
                    ZoomLevelRecommendation = 1
                End If
        End Select
End Function

'Retrieve zoom level text
Public Function ZoomLevelText(byteZL As Byte) As String
    Select Case byteZL
        Case 1
            ZoomLevelText = GetTxt(g_arrTxt, "ZOOM003")
        Case 2
            ZoomLevelText = GetTxt(g_arrTxt, "ZOOM004")
        Case 3
            ZoomLevelText = GetTxt(g_arrTxt, "ZOOM005")
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

Private Sub DrawRaceTrack()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    'Deactivate screen updating
        Application.ScreenUpdating = False
        
    'Freeze columns A-D if one of those checkboxes is activated, otherwise unfreeze
        If g_blnHorseNamesLeft Or g_blnHorseColoursLeft Or g_blnHighlightFav _
            Or (g_blnFocusedRun And g_blnHighlightFoc) Or (g_blnRaceInformation And g_blnRaceInfoWorksheet) Then
                Call basAuxiliary.Freeze(5, 0, True) 'freeze
        Else
                Call basAuxiliary.Freeze(0, 0, False) 'unfreeze
        End If

    'Formatting: row height of the different sections
        g_wksRace.Range(rows(1 + m_intTopRows), rows(5 + m_intTopRows)).EntireRow.RowHeight = 15 'above the race track
        g_wksRace.Range(rows(g_intHorsesEnrolled * 2 + 6 + 1 + m_intTopRows), _
            rows(g_intHorsesEnrolled * 2 + 52 + m_intTopRows)).EntireRow.RowHeight = 15 'below the race track
        g_wksRace.rows(g_intHorsesEnrolled * 2 + 20 + m_intTopRows).RowHeight = 20 'headline of the ranking list
    'Display race data on the top
        With g_wksRace.Cells(2 + m_intTopRows, 6)
            .Font.name = "Arial Black"
            .Value = m_strRaceName & " " & m_strRaceYear & " - " & m_strRaceTrack & ", " & m_strRaceLocation _
                    & " (" & m_strRaceCountry & ")"
        End With
        With g_wksRace.Cells(3 + m_intTopRows, 6)
            .Font.name = "Arial"
            .Font.Bold = True 'Fettschrift
            .Value = m_strRaceType & " " & GetTxt(g_arrTxt, "RACE007") & " " & _
                m_intTrackLength & GetTxt(g_arrTxt, "RACE008") & " - " & m_strTrackSurface
        End With
    'Formatting: columns on the left side of the starting grid dependent on the zoom level
        g_wksRace.Columns(1).ColumnWidth = g_byteTrackZoom 'left margin
        g_wksRace.Columns(2).ColumnWidth = g_byteTrackZoom 'horse colours
        g_wksRace.Columns(3).ColumnWidth = g_byteTrackZoom 'empty column
        g_wksRace.Columns(4).ColumnWidth = 2 + g_byteTrackZoom 'horse numbers
        g_wksRace.Columns(5).ColumnWidth = 20 + (g_byteTrackZoom * 6) 'horse names
        g_wksRace.Range(Columns(6), Columns(m_intLeftColumns + 5)).ColumnWidth = m_dblTrackCellWidth 'cell length of the race track
        g_wksRace.Columns(m_intLeftColumns - 3).ColumnWidth = 3 + g_byteTrackZoom 'starting box numbers
    'Formatting: race track
        g_wksRace.Range(Columns(m_intLeftColumns + 6), Columns(m_intTrackLength + m_intLeftColumns + 6 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor))).ColumnWidth = m_dblTrackCellWidth  'cell length of the race track
        g_wksRace.Range(rows(6 + m_intTopRows), rows(g_intHorsesEnrolled * 2 + 6 + m_intTopRows)).RowHeight = m_intTrackCellHeight   'cell width of the race track
        g_wksRace.Range(Cells(4 + m_intTopRows, 1), Cells(g_intHorsesEnrolled * 2 + 19 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor))).Interior.Color = m_lngTrackColour  'track colour
    'Formatting: font
        g_wksRace.Range(Cells(4 + m_intTopRows, 1), Cells(g_intHorsesEnrolled * 2 + 8 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor))).Font.name = "Arial"
    'Draw start boxes
        For i = 6 To (g_intHorsesEnrolled * 2 + 6) Step 2 'Für jeden Startplatz eine Box
            g_wksRace.Range(Cells(i + m_intTopRows, m_intLeftColumns + 1), Cells(i + m_intTopRows, m_intLeftColumns + 6)).Interior.ColorIndex = 1 'schwarz
        Next i
        g_wksRace.Range(Cells(6 + m_intTopRows, m_intLeftColumns + 6), Cells(g_intHorsesEnrolled * 2 + 6 + m_intTopRows, m_intLeftColumns + 6)).Interior.ColorIndex = 1 'schwarz
    'Label the start boxes
        With g_wksRace.Range(Cells(7 + m_intTopRows, m_intLeftColumns - 3), Cells(g_intHorsesEnrolled * 2 + 5 + m_intTopRows, m_intLeftColumns - 3))
            .Font.ColorIndex = 1 'colour: black
            .Font.Size = m_intFontSize
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        For i = 1 To (g_intHorsesEnrolled) 'one start box for each horse
            g_wksRace.Cells(5 + 2 * i + m_intTopRows, m_intLeftColumns - 3).Value = GetTxt(g_arrTxt, "RACE010") & " " & i
        Next i
    'Display meters above and below the race track
        For i = g_intTrackMetres To m_intTrackLength Step g_intTrackMetres
            With g_wksRace.Cells(4 + m_intTopRows, i + m_intLeftColumns)
                .Font.name = "Arial"
                .Font.Bold = True 'Fettschrift
                .Value = i & GetTxt(g_arrTxt, "RACE008") '"m"
            End With
            With g_wksRace.Cells(g_intHorsesEnrolled * 2 + 8 + m_intTopRows, i + m_intLeftColumns)
                .Font.name = "Arial"
                .Font.Bold = True 'Fettschrift
                .Value = i & GetTxt(g_arrTxt, "RACE008") '"m"
            End With
        Next i
    'Formatting: horse names on the left
        With g_wksRace.Range(Cells(6 + m_intTopRows, 4), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, 5))
            .Font.Color = m_lngTrackColour  'track colour, so that the names are not visible
            .IndentLevel = 1 'text indented
            .Font.Size = m_intFontSize
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    'Hoof prints (displayed as dots ........)
        With g_wksRace.Range(Cells(5 + m_intTopRows, m_intLeftColumns - 1), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns))
            .Font.Size = m_intFontSize
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
    'Formatting: finish area
        g_wksRace.Columns(m_intTrackLength + m_intLeftColumns + 5).ColumnWidth = m_dblTrackCellWidth 'width of the finish line
        g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 5), _
            Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 5)).Interior.ColorIndex = 56 'colour of the finish line: dark grey
        With g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 1)
            .Font.name = "Arial"
            .Font.Bold = True
            .Value = m_intTrackLength & GetTxt(g_arrTxt, "RACE008") 'distance in meters
        End With
        With g_wksRace.Cells(g_intHorsesEnrolled * 2 + 8 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 1)
            .Font.name = "Arial"
            .Font.Bold = True
            .Value = m_intTrackLength & GetTxt(g_arrTxt, "RACE008") 'distance in meters
        End With
    'Area behind the finish line
        g_wksRace.Columns(m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).ColumnWidth = 18 + (g_byteTrackZoom * 6)
        g_wksRace.Columns(m_intTrackLength + m_intLeftColumns + 8 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).ColumnWidth = g_byteTrackZoom * 3
    'Formatting: horse names in the finish
        With g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), Cells(g_intHorsesEnrolled * 2 + 7 + (2 * g_intGMPspeedfactor) + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)))
            .Font.ColorIndex = 1 'colour: black
            .IndentLevel = 1 'text indented
            .Font.Size = m_intFontSize
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    
    'Advertising below the race track
        g_wksRace.Range(rows(g_intHorsesEnrolled * 2 + 9 + m_intTopRows), _
            rows(g_intHorsesEnrolled * 2 + 19 + m_intTopRows)).EntireRow.RowHeight = m_intAdvertisingHeight 'row height according to the zoom level

        If m_strAdvertising = "Y" Then
            Dim advPos As Integer 'column position for next ad
            advPos = m_intLeftColumns + 5
            For i = 1 To UBound(m_arr_varAdv) 'draw the advertising
                z = 3 'set the pointer to the first colour code
                For j = advPos To advPos + m_wksAdv.Cells(2, m_arr_varAdv(i)) - 1
                    If j >= m_intTrackLength + m_intLeftColumns + 5 Then Exit For
                    For k = g_intHorsesEnrolled * 2 + 9 + m_intTopRows To g_intHorsesEnrolled * 2 + 19 + m_intTopRows
                        g_wksRace.Cells(k, j).Interior.Color = m_wksAdv.Cells(z, m_arr_varAdv(i)).Value
                        z = z + 1
                    Next k
                Next j
                advPos = advPos + m_wksAdv.Cells(2, m_arr_varAdv(i))
            Next i
        End If
    
    'Formatting: finish photo (headline)
        g_wksRace.Cells(2 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).Font.name = "Arial Black"  '"FOTOFINISH!"
        With g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)) '"Zielfoto..."
            .Font.Size = 14
            .Font.Bold = True
        End With
    'Formatting: finish photo and ranking list
        g_wksRace.Range(Columns(m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), Columns(m_intTrackLength + m_intLeftColumns + 168 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor))).ColumnWidth = m_dblRankingsWidth / 10   'column width according to the zoom level
        g_wksRace.Range(Columns(m_intTrackLength + m_intLeftColumns + 169 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor)), Columns(m_intTrackLength + m_intLeftColumns + 169 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor))).ColumnWidth = m_dblRankingsWidth   'column width according to the zoom level
        'Formatting: photo of the winner
        g_wksRace.Range(Columns(m_intTrackLength + m_intLeftColumns + 170 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor)), Columns(m_intTrackLength + m_intLeftColumns + 190 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor))).ColumnWidth = 2
        With g_wksRace.Range(Cells(g_intHorsesEnrolled * 2 + 20 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 170 + (2 * 10 * g_intGMPspeedfactor)), _
                        Cells(g_intHorsesEnrolled * 2 + 21 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 172 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor)))
            .Font.Size = 14
            .Font.Bold = True
        End With
    
    'Formatting: race information on the worksheet
        If g_blnRaceInformation And g_blnRaceInfoWorksheet Then
            Call basAuxiliary.RaceInfoWorksheet(g_lngRaceInfoBackColour, g_lngRaceInfoForeColour, m_intTopRows)
        End If
    
    'Place the cursor far away
        g_wksRace.Cells(100 + m_intTopRows, 1).Select
    'Activate screen updating
        Application.ScreenUpdating = True
        
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Draw horse names during the race
Private Sub DrawHorseNames()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    'Pferdenamen am linken Rand platzieren wenn Checkbox angehakt ist
        If g_blnHorseNamesLeft Then
            For i = 1 To g_intHorsesEnrolled
                If g_arr_varHorses(i, 0) = "START" Then
                    g_wksRace.Cells(g_arr_varHorses(i, 3), 4).Value = "#" & g_arr_varHorses(i, 11)  'Startnummer
                    g_wksRace.Cells(g_arr_varHorses(i, 3), 5).Value = g_arr_varHorses(i, 1)  'Name des Pferds
                End If
            Next i
        'Optimale Spaltenbreite
            g_wksRace.Range(Columns(3), Columns(4)).EntireColumn.AutoFit
        End If
    'Pferdenamen im Ziel anzeigen wenn Checkbox angehakt ist
        If g_blnHorseNamesFinish Then
            For j = 1 To g_intHorsesEnrolled
                If g_arr_varHorses(j, 0) = "START" Then
                    g_wksRace.Cells(g_arr_varHorses(j, 3), m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).Value = _
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
        .BackColor = g_lngRaceInfoBackColour
        .Caption = GetTxt(g_arrTxt, "USERFORM006")

        'Labels mit Name und Distanz des Rennens
        With .lbl_RI1
            .BackColor = g_lngRaceInfoBackColour
            .ForeColor = g_lngRaceInfoForeColour
            .Caption = m_strRaceName & " " & m_strRaceYear
            .Font.Size = 12
            .Font.Bold = True
            .AutoSize = True
        End With
        With .lbl_RI2
            .BackColor = g_lngRaceInfoBackColour
            .ForeColor = g_lngRaceInfoForeColour
            .Font.Size = 12
            .Caption = GetTxt(g_arrTxt, "RACE024") & ": " & m_intTrackLength & GetTxt(g_arrTxt, "RACE008")
            .AutoSize = True
        End With
        
        If g_blnRaceInfoProgressBar Then
            Set g_objLabel = .Controls.Add("Forms.Label.1", , True)
            With g_objLabel 'progress bar (background)
                .name = "lbl_RI3a_dyn"
                .Font.name = "Tahoma"
                .Font.Size = 8
                .left = 6
                .top = frmRaceInfo.Height - 30
                .Width = 200
                .Height = 12
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = g_lngRaceInfoForeColour
                .ForeColor = g_lngRaceInfoForeColour
                .BackColor = g_lngRaceInfoBackColour
                .TextAlign = fmTextAlignRight
                .Caption = m_intTrackLength
            End With
            Set g_objLabel = .Controls.Add("Forms.Label.1", , True)
            With g_objLabel 'progress bar (bar)
                .name = "lbl_RI3b_dyn"
                .Font.name = "Tahoma"
                .Font.Size = 8
                .left = 6
                .top = frmRaceInfo.Height - 30
                .Width = 0
                .Height = 12
                .BorderStyle = fmBorderStyleSingle
                .BorderColor = g_lngRaceInfoForeColour
                .ForeColor = g_lngRaceInfoBackColour
                .BackColor = g_lngRaceInfoForeColour
                .TextAlign = fmTextAlignLeft
            End With
            'Adjust UserForm height
            frmRaceInfo.Height = frmRaceInfo.Height + g_objLabel.Height + 6
        End If

        If g_blnRaceInfoLeader Then
            Set g_objLabel = .Controls.Add("Forms.Label.1", , True)
            With g_objLabel 'name of the leader
                .name = "lbl_RI4a_dyn"
'                        .Font.name = "Tahoma"
                .Font.Size = 12
                .left = 6
                .top = frmRaceInfo.Height - 30
                .Width = 200
                .Height = 18
                .ForeColor = g_lngRaceInfoForeColour
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
                .top = frmRaceInfo.Height - 30
                .Width = 200
                .Height = 18
                .ForeColor = g_lngRaceInfoForeColour
                .TextAlign = fmTextAlignCenter
                .Caption = ""
            End With
            'Adjust UserForm height
            frmRaceInfo.Height = frmRaceInfo.Height + g_objLabel.Height
        End If

        .Show (vbModeless) 'modeless
    End With
End Sub

'Meldung zum Rennstart
Private Sub RaceWelcome()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
''    'A MessageBox cannot handle unicode, so for example Cyrillic characters are displayed as question marks
'    MsgBox GetTxt(g_arrTxt, "RACE001") & " " & GetTxt(g_arrTxt, "RACE006") & " " & m_strRaceLocation & " (" & m_strRaceCountry & "). " & vbNewLine & vbNewLine & _
'            GetTxt(g_arrTxt, "RACE003") & " " & m_strRaceName & " " & GetTxt(g_arrTxt, "RACE007") & " " & m_intTrackLength & " " & GetTxt(g_arrTxt, "RACE009") & "." & vbNewLine & _
'            GetTxt(g_arrTxt, "RACE004") & " " & m_intHorsesStarting & " " & GetTxt(g_arrTxt, "RACE005"), , g_c_TOOL
    
    'Set the button mode
    g_strMsgButtons = "OK"
    'Assign the text for the pop-up
    g_strMsgCaption = g_c_tool
    g_strMsgText = GetTxt(g_arrTxt, "RACE001") & " " & GetTxt(g_arrTxt, "RACE006") & " " & m_strRaceLocation & " (" & m_strRaceCountry & "). " & vbNewLine & vbNewLine & _
            GetTxt(g_arrTxt, "RACE003") & ": " & m_strRaceName & " " & GetTxt(g_arrTxt, "RACE007") & " " & m_intTrackLength & " " & GetTxt(g_arrTxt, "RACE009") & "." & vbNewLine & _
            GetTxt(g_arrTxt, "RACE004") & " " & m_intHorsesStarting & " " & GetTxt(g_arrTxt, "RACE005")
    
    If g_blnSpeech Then Call SpeechOut(g_strMsgText)
    
    'Display the pop-up
    frmMsg_MultiPurpose.Show (vbModal) 'modal
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Pferde in Boxen stellen
Private Sub StartingGrid()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    Application.Wait (Now + TimeValue("0:00:02")) 'Verzögerung
    
    For i = 1 To g_intHorsesEnrolled
        If g_arr_varHorses(i, 0) = "START" Then
            g_wksRace.Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)) _
                .Interior.Color = g_arr_varHorses(i, 2)
        End If
    Next i
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Presentation of the horses with numbers and names
Private Sub RacePresentation()
    'In case of a runtime error
    On Error GoTo ERRORHANDLING
    
    Application.Wait (Now + TimeValue("0:00:02")) 'delay
    Application.DisplayCommentIndicator = xlCommentAndIndicator 'turn on cell comments
    
    'Display a comment field for each horse
    For i = 1 To g_intHorsesEnrolled
        If g_arr_varHorses(i, 0) = "START" Then
            With g_wksRace.Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4))
                If i = m_byteFavourite(1) Then 'enhance comment field
                    If g_blnFocusedRun And g_arr_varHorses(i, 11) = g_intFocusedRun Then
                        .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) _
                            & " (" & GetTxt(g_arrTxt, "RACE011") & ") >> " & GetTxt(g_arrTxt, "RACE012") 'horse number, name and (favourite) >> in focus
                    Else
                        .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) _
                            & " (" & GetTxt(g_arrTxt, "RACE011") & ")" 'horse number, name and (favourite)
                    End If
                ElseIf g_blnFocusedRun And g_arr_varHorses(i, 11) = g_intFocusedRun Then
                    .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) _
                        & " >> " & GetTxt(g_arrTxt, "RACE012") 'horse number, name and >> in focus
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
                If g_blnFocusedRun Then
                    If g_arr_varHorses(i, 11) = g_intFocusedRun Then
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
        g_strMsgCaption = m_strRaceName & " " & m_strRaceYear
        g_strMsgText = GetTxt(g_arrTxt, "RACE013") & " " & g_arr_varHorses(m_byteFavourite(1), 1) & _
                " (#" & g_arr_varHorses(m_byteFavourite(1), 11) & ")." & vbNewLine & vbNewLine & _
                GetTxt(g_arrTxt, "RACE015") & " " & g_arr_varHorses(m_byteFavourite(2), 1) & " (#" _
                & g_arr_varHorses(m_byteFavourite(2), 11) & ") " & vbNewLine & _
                GetTxt(g_arrTxt, "RACE017") & " " & g_arr_varHorses(m_byteFavourite(3), 1) & " (#" & _
                g_arr_varHorses(m_byteFavourite(3), 11) & ") " & GetTxt(g_arrTxt, "RACE018") & "."
        
        If g_blnSpeech Then Call SpeechOut(g_strMsgText)
        
        'Display the pop-up
        frmMsg_MultiPurpose.Show (vbModal) 'modal

    'Announce the focused horse
        If g_blnFocusedRun Then
            For i = 1 To UBound(g_arr_varHorses)
                If g_arr_varHorses(i, 11) = g_intFocusedRun Then
                    'Set the button mode
                    g_strMsgButtons = "OK"
                    'Assign the text for the pop-up
                    g_strMsgCaption = GetTxt(g_arrTxt, "RACEOPT026") 'Focused Run
                    g_strMsgText = GetTxt(g_arrTxt, "RACE021") & " " & g_arr_varHorses(i, 1) & " (#" & g_arr_varHorses(i, 11) & ")."

                    If g_blnSpeech Then Call SpeechOut(g_strMsgText)
        
                    'Display the pop-up
                    frmMsg_MultiPurpose.Show (vbModal) 'modal
                    Exit For
                End If
            Next i
        End If
                
    'Turn off cell comments (hide the horse names)
        Application.DisplayCommentIndicator = xlNoIndicator
        
    'Farben der Pferde am linken Rand platzieren wenn Checkbox angehakt ist
        If g_blnHorseColoursLeft Then
            For i = 1 To g_intHorsesEnrolled
                If g_arr_varHorses(i, 0) = "START" Then
                    g_wksRace.Cells(g_arr_varHorses(i, 3), 2).Interior.Color = g_arr_varHorses(i, 2) 'horse colour
                End If
            Next i
        End If
        
    'Pferdenamen und Startnummern am linken Rand zeigen wenn Checkbox angehakt ist
        g_wksRace.Range(Columns(4), Columns(5)).Font.ColorIndex = 1 'schwarz
    
    'Mark the favourite on the left if the checkbox is ticked
        If g_blnHighlightFav Then
            g_wksRace.Range(Cells(g_arr_varHorses(m_byteFavourite(1), 3), 4), Cells(g_arr_varHorses(m_byteFavourite(1), 3), 5)) _
                .Interior.Color = 192 'background color: red
            'Show horse data on the left
            g_wksRace.Cells(g_arr_varHorses(m_byteFavourite(1), 3), 2).Interior.Color = g_arr_varHorses(m_byteFavourite(1), 2) 'horse colour
            g_wksRace.Cells(g_arr_varHorses(m_byteFavourite(1), 3), 4).Value = "#" & g_arr_varHorses(m_byteFavourite(1), 11) 'horse number
            g_wksRace.Cells(g_arr_varHorses(m_byteFavourite(1), 3), 5).Value = g_arr_varHorses(m_byteFavourite(1), 1) _
                    & " (" & GetTxt(g_arrTxt, "RACE011") & ")" 'horse name
        End If
        
    'Adapt the frame around the focused horse
        If g_blnFocusedRun Then
            For i = 1 To UBound(g_arr_varHorses)
                If g_arr_varHorses(i, 11) = g_intFocusedRun Then
                    'Delete the frame around the horse
                    g_wksRace.Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
                        Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)) _
                        .Borders.LineStyle = xlLineStyleNone
                    'Highlight the focused horse on the left if the checkbox is ticked
                    If g_blnHighlightFoc Then
                        'Draw a new frame
                        g_wksRace.Range(Cells(g_arr_varHorses(i, 3), 4), Cells(g_arr_varHorses(i, 3), 5)) _
                            .BorderAround ColorIndex:=41, LineStyle:=xlDash, Weight:=xlThick
                        'Show horse data on the left
                        g_wksRace.Cells(g_arr_varHorses(i, 3), 2).Interior.Color = g_arr_varHorses(i, 2) 'horse colour
                        g_wksRace.Cells(g_arr_varHorses(i, 3), 4).Value = "#" & g_arr_varHorses(i, 11) 'horse number
                        g_wksRace.Cells(g_arr_varHorses(i, 3), 5).Value = g_wksRace.Cells(g_arr_varHorses(i, 3), 5).Value _
                                & " >> " & GetTxt(g_arrTxt, "RACE012") ' >> in focus
                    End If
                    Exit For
                End If
            Next i
        End If
    
    'Delay before the start
        Application.Wait (Now + TimeValue("0:00:04"))

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Start des Galopprennens
Private Sub RunRace()

    Dim strDistance As String '"Race distance"
    Dim strM As String '"m"
    Dim strLeader '"The current leader is"
    Dim intRefuse As Integer 'für Zufallsgenerator
    
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    m_intHorsesRunning = m_intHorsesStarting 'initial number of starters
    
    'One out of 100 refuses to run
    If g_blnIncidentRefuse Then
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
    
    'Boxennummern entfernen
        g_wksRace.Range(Cells(7 + m_intTopRows, m_intLeftColumns - 3), Cells(5 + 2 * g_intHorsesEnrolled + m_intTopRows, m_intLeftColumns - 3)).Value = ""
    'Verzögerung
        Application.Wait (Now + TimeValue("0:00:04"))
    'Boxen aufmachen
        g_wksRace.Range(Cells(6 + m_intTopRows, m_intLeftColumns + 6), Cells(g_intHorsesEnrolled * 2 + 6 + m_intTopRows, m_intLeftColumns + 6)).Interior.Color = m_lngTrackColour
    
    If g_blnSpeech Then Call SpeechOut(GetTxt(g_arrTxt, "RACE034"))
            
    'Race information data
        'Labels
        strDistance = GetTxt(g_arrTxt, "RACE024")
        strM = GetTxt(g_arrTxt, "RACE008")
        strLeader = GetTxt(g_arrTxt, "RACEINFO001")
        'Name and position of the leader
        m_intPosLeader = g_arr_varHorses(1, 4) - (m_intLeftColumns + 5) 'zero metres
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
            Debug.Print m_strRaceName & " (" & m_intTrackLength & ")"
            Debug.Print "Race start : " & Format(timeStart, "HH:MM:SS")
        #End If
        
        Do Until m_intPlace > m_intHorsesRunning 'solange noch nicht alle im Ziel sind
            'Zählvariable für den Zieleinlauf pro Schleifendurchlauf zurücksetzen
                m_intHorsesFinishing = 0
            'Neue Positionen berechnen
                For i = 1 To UBound(g_arr_varHorses)
                    'Geschwindigkeitsfaktor pro Durchlauf
                    g_arr_varHorses(i, 7) = speedKurz()
                    'Schrittweite pro Durchlauf (ungerundet)
                        If g_blnTactics0 Then
                        'Wenn Geschwindigkeit pro Rennphase konstant sein soll
                            g_arr_varHorses(i, 8) = _
                                (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + g_arr_varHorses(i, 7)) / 3
                        ElseIf g_blnTactics3 Then
                        'Wenn jedes Pferd in einem Renndrittel unterschiedlich schnell sein soll
                            'Berechnen in welchem Streckenabschnitt das Pferd ist
                            Select Case True
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 5) < m_intTrackLength * 1 / 3 'Pferd ist im 1. Renndrittel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 12)) / 4
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 5) < m_intTrackLength * 2 / 3 'Pferd ist im 2. Renndrittel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 13)) / 4
                                Case Else 'Pferd ist im 3. Renndrittel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 14)) / 4
                                End Select
                        ElseIf g_blnTactics6 Then
                            'Berechnen in welchem Streckenabschnitt das Pferd ist
                            Select Case True
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 5) < m_intTrackLength * 1 / 6 'Pferd ist im 1. Rennsechstel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 12)) / 4
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 5) < m_intTrackLength * 2 / 6 'Pferd ist im 2. Rennsechstel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 13)) / 4
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 5) < m_intTrackLength * 3 / 6 'Pferd ist im 3. Rennsechstel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 14)) / 4
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 5) < m_intTrackLength * 4 / 6 'Pferd ist im 4. Rennsechstel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 19)) / 4
                                Case (g_arr_varHorses(i, 4) - m_intLeftColumns - 5) < m_intTrackLength * 5 / 6 'Pferd ist im 5. Rennsechstel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 20)) / 4
                                Case Else 'Pferd ist im 6. Rennsechstel
                                    g_arr_varHorses(i, 8) = _
                                        (g_arr_varHorses(i, 5) + g_arr_varHorses(i, 6) + _
                                            g_arr_varHorses(i, 7) + g_arr_varHorses(i, 21)) / 4
                                End Select
                        End If

                    'Windschatten entfernen wenn Pferd drin
                        If g_blnSlipstream And g_blnSlipstreamShow And g_arr_varHorses(i, 22) > 0 Then
                            If g_arr_varHorses(i, 4) <= m_intTrackLength + m_intLeftColumns + 2 Then
                                Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9), _
                                    Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 12)).Interior.Pattern = xlSolid
                            End If
                        End If
                        
                    'Windschatten zurücksetzen
                        g_arr_varHorses(i, 22) = 0
                    
                    'Windschatten berechnen
                        If g_blnSlipstream Then 'if slipstreaming is activated
                            For k = 1 To UBound(g_arr_varHorses) 'loop through the horses
                            
                                If g_arr_varHorses(i, 15) - 1 = g_arr_varHorses(k, 15) _
                                    Or g_arr_varHorses(i, 15) + 1 = g_arr_varHorses(k, 15) Then 'one box above or below
                                        If g_arr_varHorses(i, 4) > g_arr_varHorses(k, 4) - 8 _
                                            And g_arr_varHorses(i, 4) < g_arr_varHorses(k, 4) - 4 Then
                                                If g_blnSlipstreamDouble Then
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
                        If g_blnSlipstream Then Debug.Print g_arr_varHorses(i, 1) & " (Current position " _
                            & g_arr_varHorses(i, 4) & "): " & g_arr_varHorses(i, 8) & " (before slipstream effect)"
                    #End If

                    'Schrittweite pro Durchlauf festlegen (1 oder 2 Spalten)
                        g_arr_varHorses(i, 8) = g_arr_varHorses(i, 8) + g_arr_varHorses(i, 22) / 2000 'Windschatteneffekt addieren
                        g_arr_varHorses(i, 9) = Round(g_arr_varHorses(i, 8), 0) 'Runden auf ganze Zahlen (1 oder 2)
                        
                        g_arr_varHorses(i, 9) = g_arr_varHorses(i, 9) * g_intGMPspeedfactor 'take the race speed factor into account
                   
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
                                .Interior.Color = m_lngTrackColour
                        'Neue Position des Pferds festlegen (nur wenn Pferd noch läuft)
                            If g_arr_varHorses(i, 0) = "START" Then
                                g_arr_varHorses(i, 4) = g_arr_varHorses(i, 4) + g_arr_varHorses(i, 9)
                                'RECORD.. g_arr_varHorses(i, 4) in Array schreiben, das ist neue die Position in dieser Runde
                                '   i ist die nummer im array(?)
                            End If

                        'Windschatten zeichnen
                            If g_blnSlipstream And g_blnSlipstreamShow And g_arr_varHorses(i, 22) > 0 _
                                And g_arr_varHorses(i, 4) <= m_intTrackLength + m_intLeftColumns Then
                                    Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9), _
                                        Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 12)).Interior.Pattern = xlGray16
                                'For debugging purposes
                                #If Debugging Then
                                    Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 9).Value = "*"
                                #End If
                            End If
                            
                        'Pferd neu setzen (auch die, die schon im Ziel sind wegen dem Rendering)
                            Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)) _
                                .Interior.Color = g_arr_varHorses(i, 2)

                        'Hufspuren zeichnen
                            If g_blnHoofprints Then Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 8).Value = "."
                    End If
                    
                    'Horizontal scrolling
                    If g_blnFocusedRun Then
                        'Focused Run
                        If g_arr_varHorses(i, 11) = g_intFocusedRun And g_arr_varHorses(i, 0) = "START" Then
                            If g_arr_varHorses(i, 4) > (ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.Count).Column _
                                                    - ActiveWindow.VisibleRange.Column) / 2 Then 'Focused horse is in the middle of the screen
                                    ActiveWindow.ScrollColumn = ActiveWindow.ScrollColumn + g_arr_varHorses(i, 9)
                            End If
                        End If
                    Else
                        'No Focused Run
                        If g_arr_varHorses(i, 4) > ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.Count).Column _
                                                - ((ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.Count).Column _
                                                    - ActiveWindow.VisibleRange.Column) * 1 / 10) _
                            And ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.Count).Column <= m_intTrackLength + m_intLeftColumns + 2 Then
                                'Scrollen
                                ActiveWindow.ScrollColumn = ActiveWindow.VisibleRange.Column _
                                                    + ((ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.Count).Column _
                                                    - ActiveWindow.VisibleRange.Column) * 8 / 10)
                        End If
                    End If
                    
                    If g_blnRaceInformation Then
                        'Get the position and name of the leader
                            If (g_arr_varHorses(i, 4) - m_intLeftColumns - 5) > m_intPosLeader Then
                                m_intPosLeader = g_arr_varHorses(i, 4) - m_intLeftColumns - 5
                                m_strNameLeader = g_arr_varHorses(i, 1)
                            End If
                        
                        'Refresh the race information data (in a pop-up)
                        If g_blnRaceInfoPopup Then
                            'Display the race distance progress bar
                                If g_blnRaceInfoProgressBar Then
                                    With frmRaceInfo.Controls("lbl_RI3b_dyn")
                                        .Width = CInt(200 / m_intTrackLength * m_intPosLeader)
                                        .Caption = m_intPosLeader
                                    End With
                                End If
                            'Display the name of the leader
                                If g_blnRaceInfoLeader Then
                                    If m_intPosLeader > 20 And m_intPosLeader < (m_intTrackLength - 20) Then 'display name between 20m from the start to 20m to the finish
                                        frmRaceInfo.Controls("lbl_RI4a_dyn").Caption = strLeader
                                        frmRaceInfo.Controls("lbl_RI4b_dyn").Caption = m_strNameLeader
                                    Else
                                        frmRaceInfo.Controls("lbl_RI4a_dyn").Caption = ""
                                        frmRaceInfo.Controls("lbl_RI4b_dyn").Caption = ""
                                    End If
                                End If
                        End If
    
                        'Refresh the race information data (on the worksheet)
                        If g_blnRaceInfoWorksheet Then
                            'Display the race distance progress
                                If g_blnRaceInfoProgressBar Then
                                    With g_wksRace.Cells(3 + m_intTopRows, 2)
                                        .Value = strDistance & ": " & m_intPosLeader & strM & " / " & m_intTrackLength & strM
                                    End With
                                End If
                            'Display the name of the leader
                                If g_blnRaceInfoLeader Then
                                    If m_intPosLeader > 20 And m_intPosLeader < (m_intTrackLength - 20) Then 'display name between 20m from the start to 20m to the finish
                                        With g_wksRace.Cells(1 + m_intTopRows, 2)
                                            .Value = strLeader
                                        End With
                                        With g_wksRace.Cells(2 + m_intTopRows, 3)
                                            .Value = m_strNameLeader
                                        End With
                                    Else
                                        g_wksRace.Cells(1 + m_intTopRows, 2).Value = ""
                                        g_wksRace.Cells(2 + m_intTopRows, 3).Value = ""
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
                    
                        If g_arr_varHorses(i, 4) >= m_intTrackLength + m_intLeftColumns + 5 Then 'Ziellinie erreicht
                            g_arr_varHorses(i, 0) = "BERECHNUNG"
                            m_intHorsesFinishing = m_intHorsesFinishing + 1 'zählen, wie viele Pferde in diesem Durchlauf ins Ziel kommen
                        End If
                    End If
                Next i
                
                If m_intHorsesFinishing > 0 Then
                    If m_blnWin = False Then
                        If m_intHorsesFinishing > 1 Then
                            m_blnPhotofinish = True 'Fotofinish (Zielfoto s/w)
                            'Text anpassen
                                g_wksRace.Cells(2 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).Value = GetTxt(g_arrTxt, "RACE025")  '"FOTOFINISH!"
                        End If
                        Call Zielfoto 'Zielfoto machen
                    End If
                    m_blnWin = True 'damit das Zielfoto nur 1x gemacht wird
                    Call Platzierung 'Absprung in die Platzberechnung, wenn mehr als ein Pferd in dieser Runde ins Ziel kommen
                End If
            DoEvents 'Rendering
        Loop

        'For debugging purposes: Calculate race time
        #If Debugging Then
            Debug.Print "Race finish: " & Format(Now, "HH:MM:SS")
            Debug.Print "Race time  : " & Format(Now - timeStart, "HH:MM:SS")
        #End If
        
    'Close the pop-up with race information
        If g_blnRaceInformation Then
            If g_blnRaceInfoPopup Then Unload frmRaceInfo 'close the pop-up
            If g_blnRaceInfoWorksheet Then Call basAuxiliary.RaceInfoWorksheet(xlNone, 0, m_intTopRows) 'reset: white cell, black font
        End If
    
    'In case of a photo finish
        If m_blnPhotofinish = True Then
            'Delay
                Application.Wait (Now + TimeValue("0:00:02"))
            'Unfreeze the window pane if it was frozen
                If g_blnHorseNamesLeft Or g_blnHorseColoursLeft Or g_blnHighlightFav _
                    Or (g_blnFocusedRun And g_blnHighlightFoc) Or (g_blnRaceInformation And g_blnRaceInfoWorksheet) Then Call basAuxiliary.Freeze(0, 0, False)
            'Scroll
                On Error Resume Next
                ActiveWindow.ScrollColumn = m_intTrackLength + m_intLeftColumns + 6 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)
                On Error GoTo 0
            'Black background
                g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), _
                    Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 168 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor))).Interior.ColorIndex = 1
            'Text
                g_wksRace.Cells(2 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).Value = ""
                g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).Value = GetTxt(g_arrTxt, "RACE026")  '("Photo creation")
            'Delay
                Application.Wait (Now + TimeValue("0:00:04"))
            'Show the finishing photo
                Call ShowFinishPhoto
            'Text
                g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).Value = ""
                g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).Value = GetTxt(g_arrTxt, "RACE027")  '("Photo evaluation")
            'Delay
                Application.Wait (Now + TimeValue("0:00:04"))
            'Text
                g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).Value = GetTxt(g_arrTxt, "RACE032")  '("Finishing photo")
        End If

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Platzierung berechnen wenn ein oder mehrere Pferde in einem Schleifendurchlauf ins Ziel kommen
Private Sub Platzierung()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    ReDim m_arr_varResultsCalc(1 To m_intHorsesFinishing, 0 To 4)
    
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
Private Sub Zielfoto()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
            
    'Daten eintragen
        For j = 1 To UBound(m_arr_varPhotofinish)
            m_arr_varPhotofinish(j, 0) = g_arr_varHorses(j, 3) 'Spur
            m_arr_varPhotofinish(j, 1) = g_arr_varHorses(j, 4) 'Position des Pferds
            m_arr_varPhotofinish(j, 2) = g_arr_varHorses(j, 11) 'Startnummer
        Next j
    'Blitz wenn Fotofinish
        If m_blnPhotofinish Then
            For k = 1 To 8
                With g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), _
                    Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)))
                        .Interior.ColorIndex = 1 'schwarz
                        .Interior.ColorIndex = 0 'weiß
                End With
            Next k
            g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), _
                Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor))).Interior.Color = m_lngTrackColour
        End If

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Show the photo of the finish
Private Sub ShowFinishPhoto()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    'Prepare a variable of type long, otherwise an overflow occurs in long distant races
    'when multiplying the track length for calculating the exact position
        Dim lngFin As Long
        lngFin = m_intTrackLength 'copy the track length into the long variable
    
    'Draw the background
        Application.ScreenUpdating = False 'deactivate screen updating
        'Photo frame
        g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 168 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor))) _
                .BorderAround ColorIndex:=0, Weight:=xlMedium
        If m_blnPhotofinish Then
            'Track and finish line black-and-white (in case of a tight race)
            g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 168 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor))).Interior.ColorIndex = 1  'Rasen schwarz
            g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 139 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 148 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor))).Interior.ColorIndex = 0   'Zielline weiß
            Else
            'Track and finish line in original colours
            g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 168 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor))).Interior.Color = m_lngTrackColour
            g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 139 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 148 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor))).Interior.ColorIndex = 56   'Zielline grau
        End If
    'Draw the horses
        Dim lngHorse As Long 'Horse on the photo
        For i = 1 To UBound(m_arr_varPhotofinish)
            If m_arr_varPhotofinish(i, 1) >= m_intTrackLength + m_intLeftColumns + 5 - 13 Then  'only when the horse is in the photo
                If m_arr_varPhotofinish(i, 1) - 7 >= m_intTrackLength + m_intLeftColumns + 5 - 13 Then 'completely in the photo
                    lngHorse = m_arr_varPhotofinish(i, 1) * 10 - lngFin * 9 - 80 - 17  'rear part of the horse
                Else 'partially  in the photo
                    lngHorse = m_intTrackLength + 3  'rear part of the horse
                End If
                Range(Cells(m_arr_varPhotofinish(i, 0), m_arr_varPhotofinish(i, 1) * 10 - lngFin * 9 - 1 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), _
                    Cells(m_arr_varPhotofinish(i, 0), lngHorse + 17 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor))) _
                    .Interior.Color = g_arr_varHorses(i, 2) 'draw the horse
            End If
        Next i

    'Activate screen updating
        Application.ScreenUpdating = True
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Ergebnisliste anzeigen
Private Sub ShowRankings()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    'Verzögerung
        Application.Wait (Now + TimeValue("0:00:02"))
    
    'Pop-up
        'Set the button mode
        g_strMsgButtons = "OK"
        'Assign the text for the pop-up
        g_strMsgCaption = m_strRaceName & " " & m_strRaceYear
        g_strMsgText = GetTxt(g_arrTxt, "RACE028") & vbNewLine & GetTxt(g_arrTxt, "RACE029")
        
        If g_blnSpeech Then Call SpeechOut(g_strMsgText)
        
        'Display the pop-up
        frmMsg_MultiPurpose.Show (vbModal) 'modal
        
    'Unfreeze the window pane
        If g_blnHorseNamesLeft Or g_blnHorseColoursLeft Or g_blnHighlightFav _
            Or (g_blnFocusedRun And g_blnHighlightFoc) Or (g_blnRaceInformation And g_blnRaceInfoWorksheet) Then Call basAuxiliary.Freeze(0, 0, False)
    'Scrollen zu Ergebnistafel
        Call basAuxiliary.Scroll(m_intTrackLength, m_intTopRows + (g_intHorsesEnrolled * 2 + 9))
    'Anzeigetafel
        With g_wksRace.Range(Cells(g_intHorsesEnrolled * 2 + 20 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), _
            Cells(g_intHorsesEnrolled * 2 + 20 + m_intHorsesStarting + 1 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 168 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)))
                .BorderAround ColorIndex:=0, Weight:=xlMedium 'Rahmen
                .Interior.Color = 16777215 'Hintergrund
                .Font.name = "Courier New"
                .Font.Size = 12
                .NumberFormat = "@" 'Textformat
        End With
        With g_wksRace.Cells(g_intHorsesEnrolled * 2 + 20 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor))  'Überschrift
            .Font.Size = 14
            .Font.Bold = True 'Fettschrift
            .IndentLevel = 1 'Text eingerückt
        End With
    
    'Ergebnisse eintragen
        'Überschrift
            g_wksRace.Cells(g_intHorsesEnrolled * 2 + 20 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).Value _
                = m_strRaceName & " " & m_strRaceYear & " - " & m_strRaceLocation
        'Verzögerung
            If g_blnRankingDelay Then Application.Wait (Now + TimeValue("0:00:04"))
        'Platzierungen anzeigen
            Dim intPositionName As Integer 'Position of the horse names
            intPositionName = 0
            For i = UBound(m_arr_varResults) To 1 Step -1
                If g_blnRankingColours Then
                    intPositionName = 12
                    Range(Cells(g_intHorsesEnrolled * 2 + 20 + i + m_intTopRows, m_intTrackLength + m_intLeftColumns + 12 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)), _
                        Cells(g_intHorsesEnrolled * 2 + 20 + i + m_intTopRows, m_intTrackLength + m_intLeftColumns + 23 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor))) _
                        .Interior.Color = m_arr_varResults(i, 4) 'Farbe des Pferds
                End If
                Cells(g_intHorsesEnrolled * 2 + 20 + i + m_intTopRows, m_intTrackLength + m_intLeftColumns + 15 + intPositionName + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).Value = m_arr_varResults(i, 1) & "."  'Platzierung
                Cells(g_intHorsesEnrolled * 2 + 20 + i + m_intTopRows, m_intTrackLength + m_intLeftColumns + 22 + intPositionName + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).Value = m_arr_varResults(i, 3) & _
                    " (#" & m_arr_varResults(i, 2) & ")" 'Name und Startnummer des Pferds
                'Wenn Checkbox zur Spannungssteigerung angehakt
                If g_blnRankingDelay Then Application.Wait (Now + TimeValue("0:00:01")) 'Verzögerung
            Next i

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Kurzfristige Geschwindigkeit der Pferde (Faktor wird bei jedem Schleifendurchlauf neu berechnet)
Function speedKurz() As Double
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    Randomize 'Zufallsgenerator zurücksetzen
    speedKurz = (Int((m_lngSpeedLoopHigh - m_lngSpeedLoopLow + 1) * Rnd + m_lngSpeedLoopLow) + 100000) / 100000 'Zufallszahl
    
    Exit Function
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Function

'Pferd mit Siegerkranz zeichnen
Private Sub DrawWinner()
    'In case an error occurs
    On Error GoTo ERRORHANDLING

    'Paint the horse
        Call PaintPicture("WINNER", m_intTrackLength + m_intLeftColumns + 168 + 19 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor), _
                                    m_intTopRows + g_intHorsesEnrolled * 2 + 23 + 12, _
                                    m_intTopRows + g_intHorsesEnrolled * 2 + 23, _
                                    m_intTrackLength + m_intLeftColumns + 170 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor))
        
    'Name des Siegers
        m_strWinner = ""
        For i = 1 To UBound(m_arr_varResults)
            If m_arr_varResults(i, 1) = 1 Then
                If i > 1 Then 'wenn mehrere Pferde gewinnen
                    m_strWinner = m_strWinner & " und "
                    m_blnDeadHeat = True
                    Application.Wait (Now + TimeValue("0:00:02")) 'Verzögerung
                End If
                m_strWinner = m_strWinner & UCase(m_arr_varResults(i, 3))
            End If
        Next i
        g_wksRace.Cells(g_intHorsesEnrolled * 2 + 20 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 170 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor)).Value = GetTxt(g_arrTxt, "RACE031")
        g_wksRace.Cells(g_intHorsesEnrolled * 2 + 21 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 172 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor)).Value = m_strWinner
 
        If g_blnSpeech Then Call SpeechOut(GetTxt(g_arrTxt, "RACE031"))
        If g_blnSpeech Then Call SpeechOut(m_strWinner)
        
    'In case of a dead race
        If m_blnDeadHeat Then
            'Pop-up
                'Set the button mode
                g_strMsgButtons = "OK"
                'Assign the text for the pop-up
                g_strMsgCaption = m_strRaceName & " " & m_strRaceYear
                g_strMsgText = " " & UCase(GetTxt(g_arrTxt, "RACE033")) & "!"
                'Display the pop-up
                frmMsg_Info.Show (vbModal) 'modal
        End If
        
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Analyze bet slips
Private Sub UserFormAnalyseBetSlips()
    
    Dim ID As String
    Dim nm As String
    Dim ty As String
    Dim st As Double
    Dim od As Double
    Dim bt() As Integer
    Dim payout As Boolean
    Dim noWinner As Boolean
    
    Dim totalStake As Double
    Dim totalPayout As Double
    
    If g_colBetSlips.Count > 0 Then

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
            .Caption = GetTxt(g_arrTxt, "BET039") 'Official racing result
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
                .Caption = .Caption & GetTxt(g_arrTxt, "BET040") & " " & i & ": " _
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
                .Caption = GetTxt(g_arrTxt, "BET042") 'Placed bets
            End With
    
    noWinner = True 'Reset variable
    k = 1 'lfd. nr. wettschein
             
        For i = 1 To g_colBetSlips.Count
            payout = True
            ID = g_colBetSlips(i).ID
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
                    .Caption = nm & " (" & GetTxt(g_arrTxt, "BET001") & " " & GetTxt(g_arrTxt, "ODD001") & " " & ID & ")"
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
                    .Caption = UCase(GetTxt(g_arrTxt, "BET007")) & ": " & ty 'TYPE OF BET: xxxxx
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
                    .Caption = GetTxt(g_arrTxt, "BET041") & ": " & GetHorseName(bt(j)) & " (#" & bt(j) & ")"
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
                    .Caption = GetTxt(g_arrTxt, "BET037") & ": " & Format(st, "0.00") & " " & GetTxt(g_arrTxt, "BET035") 'Stake: xx EUR
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
                    .Caption = "  " & GetTxt(g_arrTxt, "BET038") & ": " & Format(payCash, "0.00") & " " & GetTxt(g_arrTxt, "BET035") 'Pay-out: xx EUR
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
                .Caption = GetTxt(g_arrTxt, "START012") & ": " & g_colBetSlips.Count 'Number of bet slips
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
                .Caption = GetTxt(g_arrTxt, "BET043") & ": " & totalStake & " " & GetTxt(g_arrTxt, "BET035") & " / " _
                            & GetTxt(g_arrTxt, "BET044") & ": " & totalPayout & " " & GetTxt(g_arrTxt, "BET035")
            End With

        With frmBettingAnalysis
            .Caption = m_strRaceName & " " & m_strRaceYear & " | " & m_strRaceTrack & ", " & m_strRaceLocation _
                        & " (" & m_strRaceCountry & ")"
            .Width = 400 'width of the pop-up

            If g_colBetSlips.Count <= 5 Then
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
            .Show (vbModeless)  'modeless
        End With
    End If
End Sub

'Retrieve horse name from the horse number
Private Function GetHorseName(num As Integer) As String
    Dim X As Integer
    For X = 1 To UBound(g_arr_varHorses())
        If num = g_arr_varHorses(X, 11) Then Exit For
    Next X
    GetHorseName = g_arr_varHorses(X, 1)
End Function

'Warnhinweis
Private Sub ShowWarning()
    Dim strWarningMessage As String
    
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    'Compose the warning message
    strWarningMessage = strWarningMessage & GetTxt(g_arrTxt, "WARN001") & vbNewLine & GetTxt(g_arrTxt, "WARN002")

    'Pop-up
        'Set the button mode
        g_strMsgButtons = "OK"
        'Assign the text for the pop-up
        g_strMsgCaption = GetTxt(g_arrTxt, "USERFORM003")
        g_strMsgText = strWarningMessage
        'Display the pop-up
        frmMsg_Attention.Show (vbModal) 'modal
        
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
        .lblS1.Caption = m_strRaceName & " " & m_strRaceYear 'race name and year
        .lblS2.Caption = m_strRaceType & " " & GetTxt(g_arrTxt, "RACE007") & " " & m_intTrackLength & " " & GetTxt(g_arrTxt, "RACE009") ' race type and distance
        .lblS3.Caption = m_strRaceTrack & " " & GetTxt(g_arrTxt, "RACE002") & " " & m_strRaceLocation & " (" & m_strRaceCountry & ")" 'race course, loaction and country
        .lblS6.Caption = m_intHorsesStarting & " " & GetTxt(g_arrTxt, "START001") 'number of horses starting
        If m_strRealRace = "Y" Then
            .lblS10.Caption = UCase(GetTxt(g_arrTxt, "START009")) 'REAL RACE
        Else
            .lblS10.Caption = UCase(GetTxt(g_arrTxt, "START010")) 'FICTITIOUS RACE
        End If
        .lblS8.Caption = GetTxt(g_arrTxt, "START005") 'caption of the betting section
        Call NumberBetSlips 'refresh the number of bet slips
        .lblFocus.Caption = GetTxt(g_arrTxt, "START006") 'label "Focused horse"
        .cmdS1.Caption = GetTxt(g_arrTxt, "START002") 'button "Add bet slip"
        .cmdS2.Caption = GetTxt(g_arrTxt, "START003") 'button "Start race"
        With .lblS4 'track surface
            .Caption = m_strTrackSurface 'text with the surface
            .BorderStyle = fmBorderStyleSingle 'draw a border around
            .BackColor = m_lngTrackColour 'set the color according to the track colour
        End With
        .cmdS6.Caption = GetTxt(g_arrTxt, "START008") 'button "Show horse numbers and odds"
        If g_blnBettingMode Then 'height of the pop-up
            .Height = 290 'if the betting mode is enabled
        Else
            .Height = 180 'if the betting mode is disabled
        End If
        .Show (vbModal) 'modal
    End With
End Sub

Public Sub NumberBetSlips()
    frmStart.lblBet02.Caption = GetTxt(g_arrTxt, "START012") & ": " & g_colBetSlips.Count 'number of bet slips filed in
End Sub

'Name of the gambler for placing a bet
Public Sub Gambler()
    
    'Set the button mode
    g_strMsgButtons = "CancelOK"
    'Assign the text for the pop-up
    g_strMsgText = GetTxt(g_arrTxt, "BET002")
    'Display the pop-up
    With frmInp_MultiPurpose
        .Caption = m_strRaceName & " " & m_strRaceYear
        .Show (vbModal) 'modal
    End With
    'Evaluate the name of the player
    If g_strButtonPressed = "OK" And Trim(g_strPlayerName) <> "" Then _
        Call UserFormBetSlip(g_strPlayerName)

End Sub

'Information for UserForm "BetSlip"
Private Sub UserFormBetSlip(strName As String)
    With frmBetSlip
        .Caption = strName
        .lblC1 = m_strRaceTrack & " - " & m_strRaceLocation & " (" & m_strRaceCountry & ")"
        .lblC2 = m_strRaceName & " " & m_strRaceYear
        .Show (vbModal) 'modal
    End With
End Sub

'Call UserForm with odds
Public Sub odds()
    Call UserFormOdds
End Sub

'Information for UserForm "Odds"
Private Sub UserFormOdds()
    
    Dim min As Integer, max As Integer
    Dim i As Integer, j As Integer, k As Integer
    
    With frmOdds
        .Caption = m_strRaceName & " " & m_strRaceYear
        .Width = 560
        .Height = 80
        .lblO0.Caption = GetTxt(g_arrTxt, "ODD001")
        .lblO1.Caption = GetTxt(g_arrTxt, "ODD002")
        With .lblO2
            .Caption = GetTxt(g_arrTxt, "ODD003")
            .ControlTipText = GetTxt(g_arrTxt, "ODD006")
            .TextAlign = fmTextAlignRight
        End With
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

                If g_arr_varHorses(j, 0) = "START" Then
                    Set g_objLabel = frmOdds.Controls.Add("Forms.Label.1", , True) 'upper horizontal bar
                    With g_objLabel '~speed
                        .left = 290
                        .top = 27 + k * 18
                        .Height = 7
                        If m_strRealRace = "Y" Then
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
                    End With
                End If
                
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
        .Caption = GetTxt(g_arrTxt, "ODD005")
        .ControlTipText = GetTxt(g_arrTxt, "ODD008")
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
        .Caption = GetTxt(g_arrTxt, "ODD004")
        .ControlTipText = GetTxt(g_arrTxt, "ODD007")
    End With

    frmOdds.Show (vbModal) 'modal
End Sub

'Call UserForm for showing receipt
Public Sub ShowReceipt(ID As Integer)
    Call UserFormReceipt(ID)
End Sub

'Information for UserForm "Receipt"
Private Sub UserFormReceipt(ID As Integer)
    Dim bet As String, horsename As String
    Dim i As Integer, j As Integer
    Dim bt() As Integer
    bt() = g_colBetSlips(ID).bet
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
        .Caption = g_colBetSlips(ID).GamblerName
        .lblR1 = UCase(m_strRaceLocation) ' & " (" & m_strRaceCountry & ")")
        .lblR2 = UCase(m_strRaceName)
        .lblR3 = m_intHorsesStarting & " " & UCase(GetTxt(g_arrTxt, "START001"))
        .lblR4 = UCase(g_colBetSlips(ID).BType)
        .lblR5 = UCase(bet)
        .lblR6 = UCase(GetTxt(g_arrTxt, "BET036") & " " & Format(g_colBetSlips(ID).Stake, "0.00") & " " & GetTxt(g_arrTxt, "BET035"))
        .lblR7 = g_colBetSlips(ID).ID
        .Show (vbModal) 'modal
    End With
End Sub


'CALLBACKS (Excel ribbon events)
'-------------------------------

'Callback for customUI.onLoad
Private Sub AI_GaloppSimAddinInitialize(ribbon As IRibbonUI)
    Set g_RibbonGaloppSim = ribbon
    If m_wksTxt Is Nothing Then
        With ThisWorkbook
            Set m_wksTxt = .Worksheets("Txt")
            Set m_wksTxtCountry = .Worksheets("TxtCountries")
            Set m_wksTxtInfo = .Worksheets("TxtInfo")
        End With
        Call GetTextComponents
    End If
End Sub

'Callback for lblDE getLabel
Private Sub AI_GetLabel(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case "group01GALOPPSIM" 'group "Settings"
            returnedVal = GetTxt(g_arrTxt, "AI001")
            Case "btn05GALOPPSIM" 'button "Race options"
                returnedVal = GetTxt(g_arrTxt, "BTN020")
            Case "btn10GALOPPSIM" 'button "Race options"
                returnedVal = GetTxt(g_arrTxt, "BTN001")
            Case "menu02GALOPPSIM" 'menu "Language"
                returnedVal = GetTxt(g_arrTxt, "BTN002")
                Case "cb01aGALOPPSIM" 'checkbox "excel mode"
                    returnedVal = GetTxt(g_arrTxt, "EXCELOPT010")
                Case "cb01bGALOPPSIM" 'checkbox "TV mode with Excel menu strip
                    returnedVal = GetTxt(g_arrTxt, "EXCELOPT011")
                Case "cb01cGALOPPSIM" 'checkbox "TV mode (full-screen)"
                    returnedVal = GetTxt(g_arrTxt, "EXCELOPT012")
                
        Case "group02GALOPPSIM" 'group "Race"
            returnedVal = GetTxt(g_arrTxt, "AI002")
            Case "btn30GALOPPSIM" 'button "Start race"
                returnedVal = GetTxt(g_arrTxt, getCaptionStartBtn(g_blnBettingMode))
            Case "combo01InstalledRaces" 'combobox "Race selection"
                returnedVal = GetTxt(g_arrTxt, "RACE003")
            Case "btn31GALOPPSIM" 'button "Photo of the finish"
                returnedVal = GetTxt(g_arrTxt, "BTN004")
            Case "btn32GALOPPSIM" 'button "Ranking list"
                returnedVal = GetTxt(g_arrTxt, "BTN005")
            Case "btn33GALOPPSIM" 'button "Photo of the winner"
                returnedVal = GetTxt(g_arrTxt, "BTN006")
            Case "btn34GALOPPSIM" 'button "Betting analysis"
                returnedVal = GetTxt(g_arrTxt, "BTN007")
            
        Case "menu01GALOPPSIM" 'menu "Language"
            returnedVal = GetTxt(g_arrTxt, "LANGUAGE001")
            Case "btn01aGALOPPSIM" 'button "deutsch"
                returnedVal = GetTxt(g_arrTxt, "LANGUAGE002")
            Case "btn01bGALOPPSIM" 'button "english"
                returnedVal = GetTxt(g_arrTxt, "LANGUAGE003")
'                Case "btn01cGALOPPSIM" 'button "schwiizerdütsch"
'                    returnedVal = GetTxt(g_arrTxt, "LANGUAGE004")
            Case "btn01dGALOPPSIM" 'button "russian"
                returnedVal = GetTxt(g_arrTxt, "LANGUAGE005")
            Case "btn01eGALOPPSIM" 'button "bulgarian"
                returnedVal = GetTxt(g_arrTxt, "LANGUAGE006")
        Case "btn40GALOPPSIM" 'button "Info"
            returnedVal = GetTxt(g_arrTxt, "BTN009")
        Case "btn50GALOPPSIM" 'button "Warning"
            returnedVal = GetTxt(g_arrTxt, "BTN010")
        Case "btn60GALOPPSIM" 'button "Movie"
            returnedVal = GetTxt(g_arrTxt, "BTN011")
        Case "btn70GALOPPSIM" 'button "Close"
            returnedVal = GetTxt(g_arrTxt, "BTN012")
    End Select
End Sub

'Callbacks for Tooltips
Private Sub AI_GetScreentip(control As IRibbonControl, ByRef screentip)
    Select Case control.ID
        Case "btn05GALOPPSIM" 'button "Start screen"
            screentip = GetTxt(g_arrTxt, "TIP012")
        Case "btn10GALOPPSIM" 'button "Race options"
            screentip = GetTxt(g_arrTxt, "USERFORM001")
        Case "cb01aGALOPPSIM" 'checkbox "excel mode"
            screentip = GetTxt(g_arrTxt, "EXCELOPT010")
        Case "cb01bGALOPPSIM" 'checkbox "TV mode with Excel menu strip
            screentip = GetTxt(g_arrTxt, "EXCELOPT011")
        Case "cb01cGALOPPSIM" 'checkbox "TV mode (full-screen)"
            screentip = GetTxt(g_arrTxt, "EXCELOPT012")
        Case "combo01InstalledRaces" 'combobox "Race selection"
            screentip = GetTxt(g_arrTxt, "TIP024")
        Case "btn30GALOPPSIM" 'button "Start race"
            screentip = GetTxt(g_arrTxt, "TIP025")

        Case "btn40GALOPPSIM" 'button "Info"
            screentip = GetTxt(g_arrTxt, "TIP026")
        Case "btn60GALOPPSIM" 'button "Play Movie"
            screentip = GetTxt(g_arrTxt, "MOVIE003")

    End Select
End Sub

Private Sub AI_GetSupertip(control As IRibbonControl, ByRef supertip)
    Select Case control.ID
        Case "cb01aGALOPPSIM" 'checkbox "excel mode"
            supertip = GetTxt(g_arrTxt, "TIP013")
        Case "cb01bGALOPPSIM" 'checkbox "TV mode with Excel menu strip
            supertip = GetTxt(g_arrTxt, "TIP014")
        Case "cb01cGALOPPSIM" 'checkbox "TV mode (full-screen)"
            supertip = GetTxt(g_arrTxt, "TIP015")

        Case "btn40GALOPPSIM" 'button "Info"
            supertip = GetTxt(g_arrTxt, "TIP027")
        Case "btn60GALOPPSIM" 'button "Movie"
            supertip = GetTxt(g_arrTxt, "MOVIE004")

    End Select
End Sub

'Callbacks for button status
Private Sub AI_IsButtonEnabled(control As IRibbonControl, ByRef returnedVal)
    Select Case control.ID
        Case "btn31GALOPPSIM" 'button "Photo of the finish"
            returnedVal = g_blnRaceStarted
        Case "btn32GALOPPSIM" 'button "Ranking list"
            returnedVal = g_blnRaceStarted
        Case "btn33GALOPPSIM" 'button "Photo of the winner"
            returnedVal = g_blnRaceStarted
        Case "btn34GALOPPSIM" 'button "Betting analysis"
            returnedVal = g_blnRaceStarted And g_blnBetsPlaced
    End Select
End Sub

'Initialwerte der Checkboxen im Menüband (getPressed)
Public Sub AI_ExcelModeGet(control As IRibbonControl, ByRef standardwert)
    Select Case control.ID
        Case "cb01aGALOPPSIM" 'Excel mode
            standardwert = (g_strExcelMode = "normal")
        Case "cb01bGALOPPSIM" 'TV mode with Excel menu strip
            standardwert = (g_strExcelMode = "TVmenu")
        Case "cb01cGALOPPSIM" 'TV mode (full-screen)
            standardwert = (g_strExcelMode = "TVfull")
    End Select
End Sub

'Checkboxen im Menüband (onAction)
Public Sub AI_ExcelModeSet(control As IRibbonControl, pressed As Boolean)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    Select Case control.ID
        Case "cb01aGALOPPSIM" 'Excel mode
            g_strExcelMode = "normal"
            Call ResetExcelOptions
        Case "cb01bGALOPPSIM" 'TV mode with Excel menu strip
            g_strExcelMode = "TVmenu"
            Call ExcelOptionsTVmenu
        Case "cb01cGALOPPSIM" 'TV mode (full-screen)
            g_strExcelMode = "TVfull"
            Call ExcelOptionsTVmenu
    End Select
    
    g_RibbonGaloppSim.Invalidate 'refresh the status of the checkboxes
End Sub

'Ribbon button "Race options"
Private Sub AI_OptionsRace(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    frmOptionsRace.Show (vbModal) 'display UserForm (modal)
End Sub

'Startbutton im Menüband
Private Sub AI_StartRace(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    'Leave the current race?
    If g_blnRaceStarted Then
        Dim strNextRace As String
        With ThisWorkbook.Worksheets(g_strRaceSelected)
            strNextRace = .Cells(4, 2).Value & " " & _
                .Cells(5, 2).Value & " (" & _
                .Cells(12, 2).Value & "m) - " & _
                .Cells(6, 2).Value
        End With
        
'        'A MessageBox cannot handle unicode, so for example Cyrillic characters are displayed as question marks
'        If MsgBox((GetTxt(g_arrTxt, "RACE003") & " " & g_wksRace.OLEObjects("CBraces").Object.text), _
'            vbOKCancel, g_c_TOOL) = vbCancel Then Exit Sub

        'Pop-up
            'Set the button mode
            g_strMsgButtons = "CancelOK"
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_tool
            g_strMsgText = GetTxt(g_arrTxt, "RACE003") & " " & strNextRace
            'Display the pop-up
            frmMsg_MultiPurpose.Show (vbModal) 'modal
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
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    Call show_ergebnis
End Sub

'Ribbon button "Winner"
Private Sub AI_Winner(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    Call show_winnerPhoto
End Sub

Private Sub show_winnerPhoto()
    If g_blnRaceStarted Then
        g_wksRace.Activate
        Call basAuxiliary.Scroll(m_intTrackLength + m_intLeftColumns + 9 + 160 + m_intColumsAfterFinish + (2 * 10 * g_intGMPspeedfactor), m_intTopRows + (g_intHorsesEnrolled * 2 + 9))
        If g_strPlayMode = "RS" Then frmRS_navigation.Show (vbModeless) 'modeless
    End If
End Sub

Private Sub show_ergebnis()
    If g_blnRaceStarted Then
        g_wksRace.Activate
        Call basAuxiliary.Scroll(m_intTrackLength + m_intLeftColumns + 8 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor), m_intTopRows + (g_intHorsesEnrolled * 2 + 9))
        If g_strPlayMode = "RS" Then frmRS_navigation.Show (vbModeless) 'modeless
    End If
End Sub

'Zielfoto-Button im Menüband
Private Sub AI_FinishPhoto(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    Call show_foto
End Sub

Private Sub show_foto()
    If g_blnRaceStarted Then
        g_wksRace.Activate
        'Text anpassen
            g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor)).Value = GetTxt(g_arrTxt, "RACE032") '("Zielfoto")
        Call basAuxiliary.Scroll(m_intTrackLength + m_intLeftColumns + 8 + m_intColumsAfterFinish + (2 * g_intGMPspeedfactor), m_intTopRows + 1)
        Call ShowFinishPhoto
        If g_strPlayMode = "RS" Then frmRS_navigation.Show (vbModeless) 'modeless
    End If
End Sub

'Wett-Button im Menüband
Private Sub AI_Betting(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    Call show_wetten
End Sub

Private Sub show_wetten()
    If g_blnRaceStarted And g_blnBetsPlaced Then
        Call UserFormAnalyseBetSlips
    End If
End Sub

'Info-Button im Menüband
Private Sub AI_Info(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    Call show_info
End Sub

Private Sub show_info()
    
    With frmInfo
        .Caption = g_c_tool & " - " & GetTxt(g_arrTxtInfo, "INFO08")
        .lbl_info01.Caption = GetTxt(g_arrTxtInfo, "GEN01") & vbNewLine & GetTxt(g_arrTxtInfo, "GEN02")
        
        'MultiPage captions
            'MultiPage "Info"
            For i = 0 To 6
                .multiPage_info.Pages(i).Caption = GetTxt(g_arrTxtInfo, "PAGE0" & i + 1)
            Next i
                .multiPage_info.Value = 0 'set the focus on the first page
            'MultiPage "Algorithms"
            For i = 0 To 6
                .multiPage_algo.Pages(i).Caption = GetTxt(g_arrTxtInfo, "PAGEALGO0" & i + 1)
            Next i
                .multiPage_algo.MultiRow = True 'display tabs in rows without scrolling
                .multiPage_algo.Value = 0 'set the focus on the first page
        
        'Page "GaloppSim"
            With .lbl_info_galoppsim01
                .Caption = GetTxt(g_arrTxtInfo, "INFO01") & vbNewLine & vbNewLine _
                            & GetTxt(g_arrTxtInfo, "INFO02") & vbNewLine & vbNewLine _
                            & GetTxt(g_arrTxtInfo, "INFO03") & vbNewLine & vbNewLine _
                            & GetTxt(g_arrTxtInfo, "INFO04") & vbNewLine & vbNewLine _
                            & GetTxt(g_arrTxtInfo, "INFO05") & vbNewLine & vbNewLine _
                            & GetTxt(g_arrTxtInfo, "INFO06") & vbNewLine & vbNewLine _
                            & GetTxt(g_arrTxtInfo, "INFO07") & vbNewLine & vbNewLine
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
            .lbl_info_team01a.Caption = GetTxt(g_arrTxtInfo, "TEAM001")
            .lbl_info_team01b.Caption = GetTxt(g_arrTxtInfo, "TEAM002")
            .img_info_team01.ControlTipText = GetTxt(g_arrTxtInfo, "TEAM003")
            'Florian
            .lbl_info_team02a.Caption = GetTxt(g_arrTxtInfo, "TEAM004")
            .lbl_info_team02b.Caption = GetTxt(g_arrTxtInfo, "TEAM005")
            .img_info_team02.ControlTipText = GetTxt(g_arrTxtInfo, "TEAM006")
            'Paul
            .lbl_info_team03a.Caption = GetTxt(g_arrTxtInfo, "TEAM007")
            .lbl_info_team03b.Caption = GetTxt(g_arrTxtInfo, "TEAM008")
            .img_info_team03.ControlTipText = GetTxt(g_arrTxtInfo, "TEAM009")
            'Michael
            .lbl_info_team04a.Caption = GetTxt(g_arrTxtInfo, "TEAM010")
            .lbl_info_team04b.Caption = GetTxt(g_arrTxtInfo, "TEAM011")
            .img_info_team04.ControlTipText = GetTxt(g_arrTxtInfo, "TEAM012")
            'Meike
            .lbl_info_team05a.Caption = GetTxt(g_arrTxtInfo, "TEAM013")
            .lbl_info_team05b.Caption = GetTxt(g_arrTxtInfo, "TEAM014")
            .img_info_team05.ControlTipText = GetTxt(g_arrTxtInfo, "TEAM016")
            'Natalie
            .lbl_info_team06a.Caption = GetTxt(g_arrTxtInfo, "TEAM016")
            .lbl_info_team06b.Caption = GetTxt(g_arrTxtInfo, "TEAM017")
            .img_info_team06.ControlTipText = GetTxt(g_arrTxtInfo, "TEAM018")
            'Atanas
            .lbl_info_team07a.Caption = GetTxt(g_arrTxtInfo, "TEAM019")
            .lbl_info_team07b.Caption = GetTxt(g_arrTxtInfo, "TEAM020")
            .img_info_team07.ControlTipText = GetTxt(g_arrTxtInfo, "TEAM021")

            .lbl_info_team08a.Visible = False
            .lbl_info_team08b.Visible = False
            .img_info_team08.Visible = False
            
            'Vertical scrollbar
            With .multiPage_info.Pages(1)
                .ScrollBars = fmScrollBarsVertical
                .ScrollHeight = 400
                .KeepScrollBarsVisible = fmScrollBarsNone
            End With
            
        'Page "Algorithms"
            .img_info_algorithms01.ControlTipText = GetTxt(g_arrTxtInfo, "ALGO01")
            .img_info_algorithms02.ControlTipText = GetTxt(g_arrTxtInfo, "ALGO01")
            With .chk_info_algorithms01
                .Caption = GetTxt(g_arrTxtInfo, "ALGO02")
                .Font.Size = 20
                .Font.Bold = True
                .ControlTipText = GetTxt(g_arrTxtInfo, "ALGO03")
            End With
            'Algorithm 01 'overall race algorithm
                With .lbl_algo_01_00
                    .Caption = GetTxt(g_arrTxtInfo, "PAGEALGO01")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_01_01
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO10")
                    .top = frmInfo.lbl_algo_01_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_01_02
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO11")
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
                    .Caption = GetTxt(g_arrTxtInfo, "PAGEALGO02")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_02_01
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO15")
                    .top = frmInfo.lbl_algo_02_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_02_02
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO16")
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
                    .Caption = GetTxt(g_arrTxtInfo, "PAGEALGO03")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_03_01
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO20")
                    .top = frmInfo.lbl_algo_03_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_03_02
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO21")
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
                    .Caption = GetTxt(g_arrTxtInfo, "PAGEALGO04")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_04_01
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO25")
                    .top = frmInfo.lbl_algo_04_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_04_02
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO26")
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
                    .Caption = GetTxt(g_arrTxtInfo, "PAGEALGO05")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_05_01
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO30")
                    .top = frmInfo.lbl_algo_05_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_05_02
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO31")
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
                    .Caption = GetTxt(g_arrTxtInfo, "PAGEALGO06")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_06_01
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO35")
                    .top = frmInfo.lbl_algo_06_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_06_02
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO36")
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
                    .Caption = GetTxt(g_arrTxtInfo, "PAGEALGO07")
                    .top = 6
                    .left = 6
                    .Width = 330 'label width fix
                    .Font.Bold = True
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_07_01
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO40")
                    .top = frmInfo.lbl_algo_07_00.Height + 12
                    .left = 6
                    .Width = 330 'label width fix
                    .AutoSize = True 'label height depending on the content
                End With
                With .lbl_algo_07_02
                    .Caption = GetTxt(g_arrTxtInfo, "ALGO41")
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
                
        'Page "Code"
            With .lbl_info_code01
                .Caption = GetTxt(g_arrTxtInfo, "CODE01")
                .Font.Size = 12
                .WordWrap = True
                .AutoSize = True 'label height depending on the content
            End With
            .btn_info_code01.ControlTipText = GetTxt(g_arrTxtInfo, "CODE02")
            With .lbl_info_code03
                .Caption = GetTxt(g_arrTxtInfo, "CODE03")
                .Font.Size = 12
                .WordWrap = True
                .AutoSize = True 'label height depending on the content
            End With
            With .lbl_info_code04
                .Caption = GetTxt(g_arrTxtInfo, "CODE04") & vbNewLine & vbNewLine _
                            & GetTxt(g_arrTxtInfo, "CODE05") & vbNewLine & vbNewLine _
                            & GetTxt(g_arrTxtInfo, "CODE06") & vbNewLine & vbNewLine _
                            & GetTxt(g_arrTxtInfo, "CODE07")
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
                .Caption = GetTxt(g_arrTxtInfo, "CON01") & vbNewLine _
                            & GetTxt(g_arrTxtInfo, "CON02")
                .WordWrap = True
            End With
            .lbl_info_contact01b.Caption = GetTxt(g_arrTxtInfo, "CON03a")
            .lbl_info_contact01c.Caption = GetTxt(g_arrTxtInfo, "CON03b")
            With .btn_info_contact01
                .Caption = GetTxt(g_arrTxtInfo, "CON04")
                .ControlTipText = GetTxt(g_arrTxtInfo, "CON05")
                .WordWrap = True
            End With
            .btn_info_contact02.ControlTipText = GetTxt(g_arrTxtInfo, "CON06")
            With .lbl_info_contact02
                .Font.Size = 12
                .TextAlign = fmTextAlignRight
                .Caption = GetTxt(g_arrTxtInfo, "CON07")
            End With
            With .btn_info_contact03
                .Caption = GetTxt(g_arrTxtInfo, "CON08")
                .ControlTipText = GetTxt(g_arrTxtInfo, "CON09")
                .WordWrap = True
            End With
            With .btn_info_contact04
                .Caption = GetTxt(g_arrTxtInfo, "CON10")
                .ControlTipText = GetTxt(g_arrTxtInfo, "CON11")
                .WordWrap = True
            End With
            With .btn_info_contact05
                .Caption = GetTxt(g_arrTxtInfo, "CON12")
                .ControlTipText = GetTxt(g_arrTxtInfo, "CON13")
                .WordWrap = True
            End With
            With .lbl_info_contact03
                .ControlTipText = GetTxt(g_arrTxtInfo, "CON14")
                .WordWrap = True
            End With

        'Page "Donation"
            With .lbl_info_donation01
                .Font.Size = 12
                .Caption = GetTxt(g_arrTxtInfo, "DON01") & vbNewLine & vbNewLine _
                            & GetTxt(g_arrTxtInfo, "DON02")
                .AutoSize = True 'label height depending on the content
            End With
            .btn_info_donation01.ControlTipText = GetTxt(g_arrTxtInfo, "DON03")
            With .btn_info_donation02
                .Caption = GetTxt(g_arrTxtInfo, "DON04")
                .Font.Size = 24
                .ControlTipText = GetTxt(g_arrTxtInfo, "DON05")
            End With
        
        'Page "Privacy Policy"
            With .lbl_info_privacy01
                .Caption = GetTxt(g_arrTxtInfo, "PRIVACY01") & " " _
                            & GetTxt(g_arrTxtInfo, "PRIVACY02")
                .WordWrap = True
            End With
            
        .Show (vbModal) 'modal
    End With
End Sub

'Warnhinweis-Button im Menüband
Private Sub AI_Warning(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    Call show_warning
End Sub

Private Sub show_warning()
    Call AssignDataSheets
    Call GetTextComponents
    Call ShowWarning
End Sub

'Movie button (Ribbon)
Private Sub AI_Movie2017(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
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
    Call basMovie2017.PlayMovie2017(ActiveSheet.name)
End Sub

Private Sub AI_Title(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    Call TitleScreen
        
End Sub

Public Sub TitleScreen()

    'Leave the venue?
    If g_blnRaceStarted Then
        'Set the button mode
        g_strMsgButtons = "CancelOK"
        'Assign the text for the pop-up
        g_strMsgCaption = g_c_tool
        g_strMsgText = GetTxt(g_arrTxt, "WARN003")
        'Display the pop-up
        frmMsg_MultiPurpose.Show (vbModal) 'modal
        'Evaluate the return value
        If g_strButtonPressed = "CANCEL" Then Exit Sub
    End If

    Call AssignDataSheets
    Call CreateRaceSheet
    Call basAuxiliary.AI_ExcelModeStart
    Call basAuxiliary.AI_ExcelModeEnd
    
    With g_wksRace.Range(Cells(1, 1), Cells(40, 100))
        .ColumnWidth = ZoomLevelPictures()(0)
        .RowHeight = ZoomLevelPictures()(1)
    End With

    'Deactivate screen updating
        Application.ScreenUpdating = False
    
    'Paint title picture
    k = basAuxiliary.GetPictureColumn("AI_TITLE")
    m = 2 'initial row for reading the picture
        
    For i = 1 To 40
        For j = 1 To 100
            g_wksRace.Cells(i, j).Interior.Color _
                = g_wksPicData.Cells(m, k).Value
            m = m + 1 'next row on the worksheet "Pic"
        Next j
    Next i
        
    'Place the cursor far away (in the upper right corner of the screen)
        Call CursorAway

    'Activate screen updating
        Application.ScreenUpdating = True

    g_blnRaceStarted = False
        
End Sub

'Ende-Button im Menüband
Private Sub AI_Close(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    If g_blnRaceStarted Then g_blnRaceStarted = False
    
    'Blatt löschen
    Call AI_DeleteWorksheets
    Unload frmBettingAnalysis
    
    'Reset Excel ribbon
    g_RibbonGaloppSim.Invalidate

    'Reset Excel options
    Call ResetExcelOptions
    Application.ScreenUpdating = True
End Sub

'Language button "DE"
Private Sub AI_LanguageDE(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    g_strLanguage = "DE"
    Call ChangeLanguage
End Sub

'Language button "EN"
Private Sub AI_LanguageEN(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    g_strLanguage = "EN"
    Call ChangeLanguage
End Sub

'Language button "BG"
Private Sub AI_LanguageBG(control As IRibbonControl)

    'Check whether algorithms are allowed
    If g_blnStopAlgorithms Then
        If basAuxiliary.AllowAlgorithms = False Then
            g_blnStopAlgorithms = False
        Else
            Exit Sub
        End If
    End If
    
    g_strLanguage = "BG"
    Call ChangeLanguage
End Sub

Private Sub ChangeLanguage()
    Dim oleObj As OLEObject
    
    Call GetTextComponents 'Get new texts
    
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
    captionStart = basAuxiliary.getCaptionStartBtn(g_blnBettingMode)
    
    Select Case name
        'Button labels
        Case "startrace"
            g_wksRace.OLEObjects(name).Object.Caption = GetTxt(g_arrTxt, captionStart)
        Case "fotofinish"
            g_wksRace.OLEObjects(name).Object.Caption = GetTxt(g_arrTxt, "BTN004")
        Case "results"
            g_wksRace.OLEObjects(name).Object.Caption = GetTxt(g_arrTxt, "BTN005")
        Case "winner"
            g_wksRace.OLEObjects(name).Object.Caption = GetTxt(g_arrTxt, "BTN006")
        Case "bets"
            g_wksRace.OLEObjects(name).Object.Caption = GetTxt(g_arrTxt, "BTN007")
        Case "raceoptions"
            g_wksRace.OLEObjects(name).Object.Caption = GetTxt(g_arrTxt, "BTN001")
        Case "exceloptions"
            g_wksRace.OLEObjects(name).Object.Caption = GetTxt(g_arrTxt, "BTN002")
        Case "language"
            g_wksRace.OLEObjects(name).Object.Caption = GetTxt(g_arrTxt, "LANGUAGE001")
        Case "info"
            g_wksRace.OLEObjects(name).Object.Caption = GetTxt(g_arrTxt, "BTN009")
        Case "warning"
            g_wksRace.OLEObjects(name).Object.Caption = GetTxt(g_arrTxt, "BTN010")
        Case "movie2017"
            g_wksRace.OLEObjects(name).Object.Caption = GetTxt(g_arrTxt, "BTN011")
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

    Dim wksR As Worksheet
    Dim cnt As Long
    
    For Each wksR In ThisWorkbook.Worksheets
        If left(wksR.name, 5) = "race_" Then cnt = cnt + 1
        If cnt = index + 1 Then
            With wksR
                m_colRacesInstalled.Add .name
                returnedVal = .Cells(4, 2).Value & " " & _
                                .Cells(5, 2).Value & " (" & _
                                .Cells(12, 2).Value & "m) - " & .Cells(6, 2).Value
                End With
            Exit For
        End If
    Next wksR

End Sub

'Set default race
Private Sub AI_InstalledRaces_GetSelectedItemID(control As IRibbonControl, ByRef itemID As Variant)
    If g_strRaceSelected = "" Then
        g_strRaceSelected = m_colRacesInstalled(1) 'take the first race of the collection
    End If
    itemID = g_strRaceSelected
End Sub

'Ausgewähltes Rennen setzen
Private Sub AI_InstalledRaces_Click(control As IRibbonControl, ID As String, index As Integer)
    g_strRaceSelected = m_colRacesInstalled(index + 1)
End Sub

'Execute this procedure when the workbook is being closed
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
