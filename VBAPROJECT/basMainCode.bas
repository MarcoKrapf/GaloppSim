Attribute VB_Name = "basMainCode"
Option Explicit
Option Private Module

'This module contains the main code of the race simulator with most of the logic

'GALOPPSIM - Version 148.50 - April 2018
'Horse racing simulator for Microsoft Excel
'Author: Marco Matjes - excel@marco-krapf.de - https://marco-krapf.de/excel/
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
Public Const g_c_TOOL As String = "GaloppSim (v148.50)" 'Name and version of the tool
Public Const g_c_email As String = "excel@marco-krapf.de" 'Contact e-mail address
Public g_RibbonGaloppSim As IRibbonUI 'Custom ribbon
Public g_strLanguage As String 'Interface Language
Public g_strPlayMode As String ' "AI" = AddIn (.xlam) / "RS" = Run Simple (.xlsm)
Public g_colRSbuttons As Collection 'Menu buttons in the RS edition
Public g_oleComboRaces As OLEObject 'ComboBox with installed races in the RS edition
Public g_strRaceID As String 'Unique race ID
Public g_blnRaceStarted As Boolean 'Flag indicating whether a race has been started
Public g_byteTrackZoom As Byte 'Track zoom level (1=small 2=medium 3=large)
Public g_blnHorseNamesLeft As Boolean 'Flag indicating whether the names of the horses are permanently displayed at the left margin
Public g_blnHorseColoursLeft As Boolean 'Flag indicating whether the colours of the horse are displayed on the left margin
Public g_blnHorseNamesFinish As Boolean 'Flag indicating whether the names of the horses are displayed in the destination
Public g_blnHoofprints As Boolean 'Flag indicating whether hoof prints are displayed
Public g_blnTactics0 As Boolean 'If true: No racing tactics
Public g_blnTactics3 As Boolean 'If true: Racing tactics (three phases)
Public g_blnTactics6 As Boolean 'If true: Racing tactics (six phases)
Public g_blnFocusedRun As Boolean 'Flag indicating whether the Focused Run mode is active
Public g_blnHighlightFoc As Boolean 'Flag indicating whether the horse in focus is highlighted during the race
Public g_intFocusedRun As Integer 'Number of the focused horse
Public g_blnRankingDelay As Boolean 'Flag indicating whether the results are displayed on the ranking list bottom up with delay
Public g_blnRankingColours As Boolean 'Flag indicating whether the horse colours are displayed on the ranking list
Public g_blnHighlightFav As Boolean 'Flag indicating whether the favorite is highlighted during the race
Public g_blnRaceInformation As Boolean 'Flag indicating whether a pop-up with information about the race is displayed during the race
Public g_blnBettingMode As Boolean 'Flag indicating whether bettings can be placed
Public g_blnBettingAnalysis As Boolean 'Flag indicating whether a betting analysis is performed automatically after the race
Public g_blnBetsPlaced As Boolean 'If true: Bets have been placed
Public g_strRaceSelected As String 'Selected race
Public g_arr_varHorses() As Variant 'All information about the horses
Public g_colBetSlips As Collection 'List of all betting slips
Public g_intHorsesEnrolled As Integer 'number of horses registered (including horses that don´t start)
Public g_strTxt(1 To 600) As String 'Text components (general)
Public g_strTxtInfo(1 To 150) As String 'Text components (info section)
Public g_strMsgCaption As String 'Caption for the UserForm
Public g_strMsgText As String 'Text for the UserForm
Public g_strMsgButtons As String 'Buttons for the UserForm
Public g_strPlayerName As String 'Name of the player who places a bet
Public g_strButtonPressed As String 'Return value of the pressed button
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
Dim m_strTrackSurface As String 'type of the race track (turf, dirt, snow...)
Dim m_lngTrackColour As Long 'colour of the track surface
Dim m_strRaceType As String 'race type (F = flat, S = steeplechase...)
Dim m_strRandomLane As String 'lanes fix or random (F/R)
Dim m_strRandomColour As String 'horse colours fix or random (F/R)
Dim m_strRandomOdd As String 'odds fix or random (F/R)
Dim m_strAdvertising As String 'advertising (Y/N)
Dim m_intTopRows As Integer 'number of rows at the top of the worksheet (used for the menu in the RS edition)
Dim m_intLeftColumns As Integer 'number of columns left of the start boxes
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
    
        If RaceCheck(g_strRaceSelected) = False Then
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_TOOL
            g_strMsgText = g_strTxt(98)
            'Display the pop-up
            frmMsg_Attention.Show (vbModal) 'modal

            Exit Sub
        End If
        
        If g_blnBettingMode = True Then
            Set g_colBetSlips = Nothing
            Set g_colBetSlips = New Collection
        End If
    
        Call GetRaceData 'Tabellenblatt mit ausgewähltem Rennen
        Call AssignBasicValues 'Grundsätzliche Daten auslesen bzw. festlegen
        Call GetHorseData 'Daten über das Rennen einlesen
        Call UserFormSTART 'Show Start-UserForm
        
    If g_blnRaceStarted Then
        If g_strPlayMode = "AI" Then
            Call CreateRaceSheet 'Tabellenblatt "GALOPPSIM"
        End If
        
        If g_strPlayMode = "RS" Then
            Cells.Clear 'Clear the whole worksheet
            With g_wksRace.Cells(2, 2) 'Write title
                .Font.name = "Arial Black"
                .Value = g_c_TOOL
            End With
            Call RS_HideNavi 'hide the navigation area
        End If

        Call DrawRaceTrack 'Geläuf zeichnen
        Call DrawHorseNames 'Pferdenamen am Start und im Ziel wenn angehakt
        Call RaceWelcome 'Popup zu Rennbeginn
        Call StartingGrid 'Pferde in Boxen stellen
        Call RacePresentation 'Pferde vorstellen
        Call RunRace 'Rennstart
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
    End If

    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'AuswahlCheck
Private Function RaceCheck(choice As String) As Boolean
    If choice = "" Then
        RaceCheck = False
    Else
        RaceCheck = True
    End If
End Function

Public Sub RS_AddComboBox()
    'Add combobox to the navigation panel
    Call RS_addComboboxRaces("CBraces", 196, 15, 419, 22) '"name(ID)", left, top, width, height
End Sub

'Add controls in RS mode
Public Sub RS_AddCommandButtons()
    'Error handling
    On Error GoTo ERRORHANDLING
    
    'Add buttons to the navigation panel
        Set g_colRSbuttons = New Collection
        
        '"name(ID)", left, top, width, height, font-size, font:bold, _
            background-color (hex), text(xxx is the initial caption)
        Call RS_addButton("raceoptions", 15, 40, 81, 49, 11, False, &HFFFFFF, g_strTxt(252))
        Call RS_addButton("exceloptions", 99, 40, 81, 49, 11, False, &HFFFFFF, g_strTxt(253))
        Call RS_addButton("startrace", 196, 40, 81, 49, 16, True, &HC000&, g_strTxt(278))
        Call RS_addButton("fotofinish", 280, 40, 81, 49, 11, False, &HFFFFFF, g_strTxt(280))
        Call RS_addButton("results", 364, 40, 81, 49, 11, False, &HFFFFFF, g_strTxt(279))
        Call RS_addButton("winner", 448, 40, 81, 49, 11, False, &HFFFFFF, g_strTxt(286))
        Call RS_addButton("bets", 532, 40, 81, 49, 11, False, &HFFFFFF, g_strTxt(281))
        Call RS_addButton("language", 629, 40, 81, 49, 11, False, &HFFFFFF, g_strTxt(229))
        Call RS_addButton("info", 713, 40, 81, 49, 11, False, &HFFFFFF, g_strTxt(282))
        Call RS_addButton("warning", 797, 40, 81, 49, 11, False, &HFFFFFF, g_strTxt(283))
        Call RS_addButton("movie2017", 881, 40, 81, 49, 11, False, &HFFFFFF, g_strTxt(284))

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

'Click on RS button
Public Sub RS_execute_Click(name As String)
    
    On Error GoTo NORACE 'REIN
    
    Select Case name

        Case "startrace"
            If g_oleComboRaces.Object.Value = "" Then
                'Assign the text for the pop-up
                g_strMsgCaption = g_c_TOOL
                g_strMsgText = g_strTxt(98)
                'Display the pop-up
                frmMsg_Attention.Show (vbModal) 'modal
                
                Exit Sub
            Else
                'Leave the current race?
                If g_blnRaceStarted Then
                    
'                    'A MessageBox cannot handle unicode, so for example Cyrillic characters are displayed as question marks
'                    If MsgBox((g_strTxt(11) & " " & g_wksRace.OLEObjects("CBraces").Object.text), _
'                        vbOKCancel, g_c_TOOL) = vbCancel Then Exit Sub
                        
                    'Set the button mode
                    g_strMsgButtons = "CancelOK"
                    'Assign the text for the pop-up
                    g_strMsgCaption = g_c_TOOL
                    g_strMsgText = g_strTxt(11) & " " & g_wksRace.OLEObjects("CBraces").Object.text
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
            Call movie2017
    End Select
    
    Exit Sub

NORACE:
    'Assign the text for the pop-up
    g_strMsgCaption = g_c_TOOL
    g_strMsgText = g_strTxt(98)
    'Display the pop-up
    frmMsg_Attention.Show (vbModal) 'modal
End Sub

Private Sub ShowNewRaceScreen()
    With g_wksRace.UsedRange
        .ColumnWidth = ZoomLevelPictures()(0)
        .RowHeight = ZoomLevelPictures()(1)
        .Clear
    End With
    If g_strPlayMode = "RS" Then Call RS_HideNavi 'hide the navigation area
    ActiveWindow.ScrollColumn = 1 'scroll to the left
    Call PaintPicture("NEWRACE", 100, 40, 1, 1) 'paint title picture
    'Place the cursor far away (in the upper right corner of the screen)
        g_wksRace.Cells(1, ActiveWindow.VisibleRange.Columns.Count - 1).Activate
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
        g_wksRace.Cells(2, 2).Value = g_c_TOOL
    'Paint title picture
        Call PaintPicture("RUNSIMPLE", 100, 40, 1, 1)
    'Place the cursor far away (in the upper right corner of the screen)
        g_wksRace.Cells(1, ActiveWindow.VisibleRange.Columns.Count - 1).Activate
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
                    g_strMsgText = g_strTxt(375) & vbNewLine & vbNewLine & _
                                    g_strTxt(376) & ": " & ZoomLevelText(g_byteTrackZoom) & vbNewLine & _
                                    g_strTxt(377) & ": " & ZoomLevelText(byteOpt) & vbNewLine & vbNewLine & _
                                    g_strTxt(378)
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
                m_strTrackSurface = g_strTxt(4)
            Case "D" 'dirt
                m_strTrackSurface = g_strTxt(5)
            Case "S" 'snow
                m_strTrackSurface = g_strTxt(6)
        End Select
        Select Case .Cells(11, 2).Value 'race type
            Case "F" 'flat race
                m_strRaceType = g_strTxt(170)
            Case "S" 'steeplechase
                m_strRaceType = g_strTxt(171)
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
    Select Case g_strLanguage
        Case "DE"
            j = 2
        Case "EN"
            j = 3
        Case "BG"
            j = 4
        Case "CH"
            j = 5
    End Select
    
    'Read the text components into an array
        'General text components
        For i = 1 To UBound(g_strTxt)
            g_strTxt(i) = m_wksTxt.Cells(i, j).Value
        Next i
        'Texts for the info section
        For i = 1 To UBound(g_strTxtInfo)
            g_strTxtInfo(i) = m_wksTxtInfo.Cells(i, j).Value
        Next i
    
End Sub

'Daten über die Pferde
Private Sub GetHorseData()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    'Anzahl der Starter aus Tabellenblatt auslesen
        m_intHorsesStarting = Application.WorksheetFunction.CountIf(g_wksRaceData.Columns(6), "START")
    'Arrays anlegen
    ReDim g_arr_varHorses(1 To g_intHorsesEnrolled, 0 To 21) 'Alle Daten der Pferde
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
        g_arr_varHorses(i, 0) = g_wksRaceData.Cells(1 + i, 6).Value 'Status (START, CANCELLED...)
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
                                                        '(linear von 1,50010 bis 1,49990)
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
                        * (150012 - (g_arr_varHorses(i, 6) * 100000)) * 10) / 5) * 5 'complicated formula
                Loop Until g_arr_varHorses(i, 17) >= 15 'minimum value
            End If
        'Schätzfehler für Balkenanzeige bei Wetten (+/-50)
            Randomize 'Zufallsgenerator zurücksetzen
            g_arr_varHorses(i, 18) = (Int((100 - 0 + 1) * Rnd + 0)) - 50 'random number between -50 and +50
    Next i
    
    'Favoriten errmitteln aus Grundgeschwindigkeit und Form
        'Clear the entire array
            Erase m_dblFavCalc
            
'            'Alternatively: clear the array fields one by one
'            m_dblFavCalc(1) = 0
'            m_dblFavCalc(2) = 0
'            m_dblFavCalc(3) = 0

'            'Alternatively: clear the array fields using a loop
'            For i = 1 To 3
'                 m_dblFavCalc(i) = 0
'            Next i
            
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
'    g_arr_varHorses(m_intFavourite1, 16) = 1
'    g_arr_varHorses(m_intFavourite2, 16) = 2
'    g_arr_varHorses(m_intFavourite3, 16) = 3
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

Public Sub BigPicExcelOptions()
    Call SetExcelOptions(False, False, False, _
                             False, False, False, _
                             False, True, True)
End Sub

Private Sub SetExcelOptions(blnGrid As Boolean, blnHead As Boolean, blnFormula As Boolean, _
                            blnStatus As Boolean, blnVScroll As Boolean, blnHScroll As Boolean, _
                            blnTabs As Boolean, blnFull As Boolean, blnMax As Boolean)

    With Application
        'Since some parameters depend on each other, the order of execution is important
        If g_strPlayMode = "RS" Then
            .DisplayFullScreen = blnFull 'Excel ribbon
        End If
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
            ZoomLevelText = g_strTxt(371)
        Case 2
            ZoomLevelText = g_strTxt(372)
        Case 3
            ZoomLevelText = g_strTxt(373)
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
            Or (g_blnFocusedRun And g_blnHighlightFoc) Then
                Call basAuxiliary.Freeze(5, 0, True) 'freeze
        Else
                Call basAuxiliary.Freeze(0, 0, False) 'unfreeze
        End If

    'Formatting: row height of the different sections
        g_wksRace.Range(rows(1 + m_intTopRows), rows(5 + m_intTopRows)).EntireRow.RowHeight = 15 'above the race track
        g_wksRace.Range(rows(g_intHorsesEnrolled * 2 + 6 + 1 + m_intTopRows), _
            rows(g_intHorsesEnrolled * 2 + 52 + m_intTopRows)).EntireRow.RowHeight = 15 'below the race track (Todo: adv???)
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
            .Value = m_strRaceType & " " & g_strTxt(1) & " " & _
                m_intTrackLength & g_strTxt(2) & " - " & m_strTrackSurface
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
        g_wksRace.Range(Columns(m_intLeftColumns + 6), Columns(m_intTrackLength + m_intLeftColumns + 6)).ColumnWidth = m_dblTrackCellWidth 'cell length of the race track
        g_wksRace.Range(rows(6 + m_intTopRows), rows(g_intHorsesEnrolled * 2 + 6 + m_intTopRows)).RowHeight = m_intTrackCellHeight 'cell width of the race track
        g_wksRace.Range(Cells(4 + m_intTopRows, 1), Cells(g_intHorsesEnrolled * 2 + 19 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7)).Interior.Color = m_lngTrackColour 'track colour
    'Formatting: font
        g_wksRace.Range(Cells(4 + m_intTopRows, 1), Cells(g_intHorsesEnrolled * 2 + 8 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7)).Font.name = "Arial"
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
            g_wksRace.Cells(5 + 2 * i + m_intTopRows, m_intLeftColumns - 3).Value = g_strTxt(8) & i
        Next i
    'Display meters above and below the race track
        For i = 250 To m_intTrackLength - 250 Step 250
'            g_wksRace.Range(Cells(5, i + m_intLeftColumns + 5), Cells(45, i + m_intLeftColumns + 5)).Interior.ColorIndex = 1 'comment in for development purpose
            With g_wksRace.Cells(4 + m_intTopRows, i + m_intLeftColumns)
                .Font.name = "Arial"
                .Font.Bold = True 'Fettschrift
                .Value = i & g_strTxt(2) '"m"
            End With
            With g_wksRace.Cells(g_intHorsesEnrolled * 2 + 8 + m_intTopRows, i + m_intLeftColumns)
                .Font.name = "Arial"
                .Font.Bold = True 'Fettschrift
                .Value = i & g_strTxt(2) '"m"
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
            .Value = m_intTrackLength & g_strTxt(2) 'distance in meters
        End With
        With g_wksRace.Cells(g_intHorsesEnrolled * 2 + 8 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 1)
            .Font.name = "Arial"
            .Font.Bold = True
            .Value = m_intTrackLength & g_strTxt(2) 'distance in meters
        End With
    'Area behind the finish line
        g_wksRace.Columns(m_intTrackLength + m_intLeftColumns + 7).ColumnWidth = 18 + (g_byteTrackZoom * 6)
        g_wksRace.Columns(m_intTrackLength + m_intLeftColumns + 8).ColumnWidth = g_byteTrackZoom * 3
    'Formatting: horse names in the finish
        With g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7))
            .Font.ColorIndex = 1 'colour: black
            .IndentLevel = 1 'text indented
            .Font.Size = m_intFontSize
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    
    'Advertising below the race track
        If m_strAdvertising = "Y" Then
            g_wksRace.Range(rows(g_intHorsesEnrolled * 2 + 9 + m_intTopRows), _
                rows(g_intHorsesEnrolled * 2 + 19 + m_intTopRows)).EntireRow.RowHeight = m_intAdvertisingHeight 'row height according to the zoom level
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
        g_wksRace.Cells(2 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7).Font.name = "Arial Black" '"FOTOFINISH!"
        With g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9) '"Zielfoto..."
            .Font.Size = 14
            .Font.Bold = True
        End With
    'Formatting: finish photo and ranking list
        g_wksRace.Range(Columns(m_intTrackLength + m_intLeftColumns + 9), Columns(m_intTrackLength + m_intLeftColumns + 168)).ColumnWidth = m_dblRankingsWidth / 10 'column width according to the zoom level
        g_wksRace.Range(Columns(m_intTrackLength + m_intLeftColumns + 169), Columns(m_intTrackLength + m_intLeftColumns + 169)).ColumnWidth = m_dblRankingsWidth 'column width according to the zoom level
        'Formatting: photo of the winner
        g_wksRace.Range(Columns(m_intTrackLength + m_intLeftColumns + 170), Columns(m_intTrackLength + m_intLeftColumns + 190)).ColumnWidth = 2
        With g_wksRace.Range(Cells(g_intHorsesEnrolled * 2 + 20 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 170), _
                        Cells(g_intHorsesEnrolled * 2 + 21 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 172))
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
                    g_wksRace.Cells(g_arr_varHorses(j, 3), m_intTrackLength + m_intLeftColumns + 7).Value = _
                        g_arr_varHorses(j, 1) & " (#" & g_arr_varHorses(j, 11) & ")"
                End If
            Next j
        End If
        
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Meldung zum Rennstart
Private Sub RaceWelcome()
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
''    'A MessageBox cannot handle unicode, so for example Cyrillic characters are displayed as question marks
'    MsgBox g_strTxt(9) & " " & g_strTxt(14) & " " & m_strRaceLocation & " (" & m_strRaceCountry & "). " & vbNewLine & vbNewLine & _
'            g_strTxt(11) & " " & m_strRaceName & " " & g_strTxt(1) & " " & m_intTrackLength & " " & g_strTxt(3) & "." & vbNewLine & _
'            g_strTxt(12) & " " & m_intHorsesStarting & " " & g_strTxt(13), , g_c_TOOL
    
    'Set the button mode
    g_strMsgButtons = "OK"
    'Assign the text for the pop-up
    g_strMsgCaption = g_c_TOOL
    g_strMsgText = g_strTxt(9) & " " & g_strTxt(14) & " " & m_strRaceLocation & " (" & m_strRaceCountry & "). " & vbNewLine & vbNewLine & _
            g_strTxt(11) & " " & m_strRaceName & " " & g_strTxt(1) & " " & m_intTrackLength & " " & g_strTxt(3) & "." & vbNewLine & _
            g_strTxt(12) & " " & m_intHorsesStarting & " " & g_strTxt(13)
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
                .AddComment text:="#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1) 'horse number and name
                .Comment.Shape.TextFrame.Characters.Font.Size = m_intFontSize 'font size according to the zoom level
                .Comment.Shape.TextFrame.AutoSize = True 'optimise the size of the comment field
            End With
            If i = m_byteFavourite(1) Then
                g_wksRace.Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)) _
                    .Comment.Shape.Fill.ForeColor.RGB = RGB(192, 0, 0)
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
        g_strMsgText = g_strTxt(15) & " " & g_arr_varHorses(m_byteFavourite(1), 1) & _
                " (#" & g_arr_varHorses(m_byteFavourite(1), 11) & ")." & vbNewLine & vbNewLine & _
                g_strTxt(17) & " " & g_arr_varHorses(m_byteFavourite(2), 1) & " (#" _
                & g_arr_varHorses(m_byteFavourite(2), 11) & ") " & vbNewLine & _
                g_strTxt(19) & " " & g_arr_varHorses(m_byteFavourite(3), 1) & " (#" & _
                g_arr_varHorses(m_byteFavourite(3), 11) & ") " & g_strTxt(20) & "."
        'Display the pop-up
        frmMsg_MultiPurpose.Show (vbModal) 'modal

    'Announce the focused horse
        If g_blnFocusedRun Then
            For i = 1 To UBound(g_arr_varHorses)
                If g_arr_varHorses(i, 11) = g_intFocusedRun Then
                    'Set the button mode
                    g_strMsgButtons = "OK"
                    'Assign the text for the pop-up
                    g_strMsgCaption = g_strTxt(250) 'Focused Run
                    g_strMsgText = g_strTxt(87) & " " & g_arr_varHorses(i, 1) & " (#" & g_arr_varHorses(i, 11) & ")."
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
                    & " (" & g_strTxt(390) & ")" 'horse name
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
                        g_wksRace.Cells(g_arr_varHorses(i, 3), 5).Value = g_arr_varHorses(i, 1) _
                                & " (" & g_strTxt(391) & ")" 'horse name
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
    'In case an error occurs
    On Error GoTo ERRORHANDLING
    
    'Boxennummern entfernen
        g_wksRace.Range(Cells(7 + m_intTopRows, m_intLeftColumns - 3), Cells(5 + 2 * g_intHorsesEnrolled + m_intTopRows, m_intLeftColumns - 3)).Value = ""
    'Verzögerung
        Application.Wait (Now + TimeValue("0:00:04"))
    'Boxen aufmachen
        g_wksRace.Range(Cells(6 + m_intTopRows, m_intLeftColumns + 6), Cells(g_intHorsesEnrolled * 2 + 6 + m_intTopRows, m_intLeftColumns + 6)).Interior.Color = m_lngTrackColour
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
      
    'Show pop-up with race info if checked
        If g_blnRaceInformation Then
            With frmRaceInfo
                'Place the pop-up in the upper left corner
                    .StartUpPosition = 0
                    .top = ActiveWindow.top + 20
                    .left = ActiveWindow.left + 20
                .BackColor = &H8080FF     'red (hexadezimal?) ToDo... auslagern in generelle variable
                .Caption = g_strTxt(356)
                With .lbl_RI1
                    .BackColor = &H8080FF     'red
                    .Caption = m_strRaceName & " " & m_strRaceYear
                    .Font.Bold = True
                    .AutoSize = True
                End With
                With .lbl_RI2
                    .BackColor = &H8080FF     'red
                    .Caption = g_strTxt(89) & m_intTrackLength & g_strTxt(2)
                    .AutoSize = True
                End With
                .Show (vbModeless) 'modeless
            End With
        End If
    
    'Rennen läuft
        Do Until m_intPlace > m_intHorsesStarting 'solange noch nicht alle im Ziel sind
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
                    'Schrittweite pro Durchlauf (1 oder 2 Spalten)
                    g_arr_varHorses(i, 9) = Round(g_arr_varHorses(i, 8), 0)
                Next i
            'Pferde laufen
                For i = 1 To UBound(g_arr_varHorses)
                    If g_arr_varHorses(i, 0) <> "CANCELLED" Then 'nur wenn Pferd am Start ist
                        'Pferd löschen
                            Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)) _
                                .Interior.Color = m_lngTrackColour
                        'Neue Position des Pferds festlegen (nur wenn Pferd noch läuft)
                            If g_arr_varHorses(i, 0) = "START" Then
                                g_arr_varHorses(i, 4) = g_arr_varHorses(i, 4) + g_arr_varHorses(i, 9)
                            End If
                        'Pferd neu setzen (auch die, die schon im Ziel sind wegen dem Rendering)
                            Range(Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4)), _
                                Cells(g_arr_varHorses(i, 3), g_arr_varHorses(i, 4) - 7)) _
                                .Interior.Color = g_arr_varHorses(i, 2)
                        'Wenn Checkbox angehakt ist: g_blnHoofprints zeichnen
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
                Next i
                
            'Check ob ein Pferd im Ziel ist
                For i = 1 To UBound(g_arr_varHorses)
                    If g_arr_varHorses(i, 0) = "START" Then 'nur wenn Pferd noch läuft
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
                                g_wksRace.Cells(2 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7).Value = g_strTxt(21) '"FOTOFINISH!"
                        End If
                        Call Zielfoto 'Zielfoto machen
                    End If
                    m_blnWin = True 'damit das Zielfoto nur 1x gemacht wird
                    Call Platzierung 'Absprung in die Platzberechnung, wenn mehr als ein Pferd in dieser Runde ins Ziel kommen
                End If
            DoEvents 'Rendering
        Loop
        
    'Close the pop-up with race information
        Unload frmRaceInfo
    
    'In case of a photo finish
        If m_blnPhotofinish = True Then
            'Delay
                Application.Wait (Now + TimeValue("0:00:02"))
            'Unfreeze the window pane if it was frozen
                If g_blnHorseNamesLeft Or g_blnHorseColoursLeft Or g_blnHighlightFav _
                    Or (g_blnFocusedRun And g_blnHighlightFoc) Then Call basAuxiliary.Freeze(0, 0, False)
            'Scroll
                On Error Resume Next
                ActiveWindow.ScrollColumn = m_intTrackLength
                On Error GoTo 0
            'Text
                g_wksRace.Cells(2 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7).Value = ""
                g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9).Value = g_strTxt(22) '("Photo creation")
            'Delay
                Application.Wait (Now + TimeValue("0:00:02"))
            'Show the finishing photo
                Call ShowFinishPhoto
            'Delay
                Application.Wait (Now + TimeValue("0:00:02"))
            'Text
                g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9).Value = ""
                g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9).Value = g_strTxt(23) '("Photo evaluation")
            'Delay
                Application.Wait (Now + TimeValue("0:00:02"))
            'Text
                g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9).Value = g_strTxt(28) '("Finishing photo")
            'Delay
                Application.Wait (Now + TimeValue("0:00:02"))
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
                With g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7), _
                    Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7))
                        .Interior.ColorIndex = 1 'schwarz
                        .Interior.ColorIndex = 0 'weiß
                End With
            Next k
            g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7), _
                Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 7)).Interior.Color = m_lngTrackColour
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
        g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 168)) _
                .BorderAround ColorIndex:=0, Weight:=xlMedium
        If m_blnPhotofinish Then
            'Track and finish line black-and-white (in case of a tight race)
            g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 168)).Interior.ColorIndex = 1 'Rasen schwarz
            g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 139), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 148)).Interior.ColorIndex = 0 'Zielline weiß
            Else
            'Track and finish line in original colours
            g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 168)).Interior.Color = m_lngTrackColour
            g_wksRace.Range(Cells(5 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 139), Cells(g_intHorsesEnrolled * 2 + 7 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 148)).Interior.ColorIndex = 56 'Zielline grau
        End If
    'Draw the horses
        Dim lngHorse As Long 'Horse on the photo
        For i = 1 To UBound(m_arr_varPhotofinish)
            If m_arr_varPhotofinish(i, 1) >= m_intTrackLength + m_intLeftColumns + 5 - 13 Then 'only when the horse is in the photo
                If m_arr_varPhotofinish(i, 1) - 7 >= m_intTrackLength + m_intLeftColumns + 5 - 13 Then 'completely in the photo
                    lngHorse = m_arr_varPhotofinish(i, 1) * 10 - lngFin * 9 - 80 - 17 'rear part of the horse
                Else 'partially  in the photo
                    lngHorse = m_intTrackLength + 3 'rear part of the horse
                End If
                Range(Cells(m_arr_varPhotofinish(i, 0), m_arr_varPhotofinish(i, 1) * 10 - lngFin * 9 - 1), _
                    Cells(m_arr_varPhotofinish(i, 0), lngHorse + 17)) _
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
        g_strMsgText = g_strTxt(24) & vbNewLine & g_strTxt(25)
        'Display the pop-up
        frmMsg_MultiPurpose.Show (vbModal) 'modal
        
    'Unfreeze the window pane
        If g_blnHorseNamesLeft Or g_blnHorseColoursLeft Or g_blnHighlightFav _
            Or (g_blnFocusedRun And g_blnHighlightFoc) Then Call basAuxiliary.Freeze(0, 0, False)
    'Scrollen zu Ergebnistafel
        Call basAuxiliary.Scroll(m_intTrackLength, m_intTopRows + (g_intHorsesEnrolled * 2 + 9))
    'Anzeigetafel
        With g_wksRace.Range(Cells(g_intHorsesEnrolled * 2 + 20 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9), _
            Cells(g_intHorsesEnrolled * 2 + 20 + m_intHorsesStarting + 1 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 168))
                .BorderAround ColorIndex:=0, Weight:=xlMedium 'Rahmen
                .Interior.Color = 16777215 'Hintergrund
                .Font.name = "Courier New"
                .Font.Size = 12
                .NumberFormat = "@" 'Textformat
        End With
        With g_wksRace.Cells(g_intHorsesEnrolled * 2 + 20 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9) 'Überschrift
            .Font.Size = 14
            .Font.Bold = True 'Fettschrift
            .IndentLevel = 1 'Text eingerückt
        End With
    
    'Ergebnisse eintragen
        'Überschrift
            g_wksRace.Cells(g_intHorsesEnrolled * 2 + 20 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9).Value _
                = m_strRaceName & " " & m_strRaceYear & " - " & m_strRaceLocation
        'Verzögerung
            If g_blnRankingDelay Then Application.Wait (Now + TimeValue("0:00:04"))
        'Platzierungen anzeigen
            Dim intPositionName As Integer 'Position of the horse names
            intPositionName = 0
            For i = UBound(m_arr_varResults) To 1 Step -1
                If g_blnRankingColours Then
                    intPositionName = 12
                    Range(Cells(g_intHorsesEnrolled * 2 + 20 + i + m_intTopRows, m_intTrackLength + m_intLeftColumns + 12), _
                        Cells(g_intHorsesEnrolled * 2 + 20 + i + m_intTopRows, m_intTrackLength + m_intLeftColumns + 23)) _
                        .Interior.Color = m_arr_varResults(i, 4) 'Farbe des Pferds
                End If
                Cells(g_intHorsesEnrolled * 2 + 20 + i + m_intTopRows, m_intTrackLength + m_intLeftColumns + 15 + intPositionName).Value = m_arr_varResults(i, 1) & "." 'Platzierung
                Cells(g_intHorsesEnrolled * 2 + 20 + i + m_intTopRows, m_intTrackLength + m_intLeftColumns + 22 + intPositionName).Value = m_arr_varResults(i, 3) & _
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
        Call PaintPicture("WINNER", m_intTrackLength + m_intLeftColumns + 168 + 19, _
                                    m_intTopRows + g_intHorsesEnrolled * 2 + 23 + 12, _
                                    m_intTopRows + g_intHorsesEnrolled * 2 + 23, _
                                    m_intTrackLength + m_intLeftColumns + 170)
        
    'Name des Siegers
        m_strWinner = ""
        For i = 1 To UBound(m_arr_varResults)
            If m_arr_varResults(i, 1) = 1 Then
                If i > 1 Then 'wenn mehrere Pferde gewinnen
                    m_strWinner = m_strWinner & " und "
                    m_blnDeadHeat = True
                End If
                m_strWinner = m_strWinner & UCase(m_arr_varResults(i, 3))
            End If
        Next i
        g_wksRace.Cells(g_intHorsesEnrolled * 2 + 20 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 170).Value = g_strTxt(27)
        g_wksRace.Cells(g_intHorsesEnrolled * 2 + 21 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 172).Value = m_strWinner
    
    'In case of a dead m_strRaceName
        If m_blnDeadHeat Then
            'Pop-up
                'Set the button mode
                g_strMsgButtons = "OK"
                'Assign the text for the pop-up
                g_strMsgCaption = m_strRaceName & " " & m_strRaceYear
                g_strMsgText = UCase(g_strTxt(29)) & "!"
                'Display the pop-up
                frmMsg_Info.Show (vbModal) 'modal
        End If
    
    Exit Sub
ERRORHANDLING:
    Call basAuxiliary.CodeCrash
End Sub

'Analyze bet slips
Private Sub UserFormAnalyseBetSlips()

    Dim objLabel1 As MSForms.Label '"Official racing result"
    Dim objLabel2 As MSForms.Label 'racing result (place 1-4)
    Dim objLabel3 As MSForms.Label 'name and bet slip ID
    Dim objLabel4 As MSForms.Label 'type of bet
    Dim objLabel5 As MSForms.Label 'guess
    Dim objLabel6 As MSForms.Label 'stake
    Dim objLabel7 As MSForms.Label 'payout
    
    Dim id As String
    Dim nm As String
    Dim ty As String
    Dim st As Double
    Dim od As Double
    Dim bt() As Integer
    Dim payout As Boolean
    Dim noWinner As Boolean
    
    If g_colBetSlips.Count > 0 Then

        'Headline: "Official racing result"
        Set objLabel1 = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
        With objLabel1
            With .Font
                .name = "Tahoma"
                .Size = 12
                .Bold = True
            End With
            .left = 15
            .top = 10
            .Width = 300
            .TextAlign = fmTextAlignLeft
            .Caption = g_strTxt(190) 'Official racing result
        End With

        'Race result (place 1-4)
        Set objLabel2 = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
        With objLabel2
            .Font.name = "Tahoma"
            .Font.Size = 10
            .left = 80
            .top = 30
            .Width = 300
            .Height = 50
            .TextAlign = fmTextAlignLeft
            For i = 1 To 4 'compile the label with the horses on place 1-4
                .Caption = .Caption & g_strTxt(191) & " " & i & ": " _
                                & m_arr_varResults(i, 3) & " (#" & m_arr_varResults(i, 2) & ")" & vbNewLine
            Next i
        End With

        'Headline: "Placed bets"
        Set objLabel1 = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
        With objLabel1
            With .Font
                .name = "Tahoma"
                .Size = 12
                .Bold = True
            End With
            .left = 15
            .top = 90
            .Width = 300
            .TextAlign = fmTextAlignLeft
            .Caption = g_strTxt(193) 'Placed bets
        End With
    
    noWinner = True 'Reset variable
    k = 1 'lfd. nr. wettschein
             
        For i = 1 To g_colBetSlips.Count
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
                payColor = &H8000000F 'grey ToDO... was ist das für ein Format?? --> Buch
            Else
                payCash = st / 10 * od
                payColor = 52377 'green
            End If
            
            'Write data of each bet slip to the userform
                'Name and bet slip ID
                Set objLabel3 = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
                With objLabel3
                    With .Font
                        .name = "Tahoma"
                        .Size = 10
                        .Bold = True
                    End With
                    .left = 40
                    .top = 105 + k * 12
                    .Width = 350
                    .TextAlign = fmTextAlignLeft
                    .Caption = nm & " (" & g_strTxt(110) & " " & g_strTxt(160) & " " & id & ")"
                End With
                
                k = k + 1
                
                'Type of bet
                Set objLabel4 = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
                With objLabel4
                    .Font.name = "Tahoma"
                    .Font.Size = 10
                    .left = 80
                    .top = 105 + k * 12
                    .Width = 200
                    .TextAlign = fmTextAlignLeft
                    .Caption = UCase(g_strTxt(116)) & ": " & ty 'TYPE OF BET: xxxxx
                End With
            
                k = k + 1
                
            For j = 1 To UBound(bt)
                'Guess
                Set objLabel5 = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
                With objLabel5
                    .Font.name = "Tahoma"
                    .Font.Size = 10
                    .left = 80
                    .top = 105 + k * 12
                    .Width = 200
                    .TextAlign = fmTextAlignLeft
                    .Caption = g_strTxt(192) & ": " & GetHorseName(bt(j)) & " (#" & bt(j) & ")"
                End With
                
                k = k + 1
                
            Next j
            
            'Stake
                Set objLabel6 = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
                With objLabel6
                    .Font.name = "Tahoma"
                    .Font.Size = 10
                    .left = 80
                    .top = 105 + k * 12
                    .Width = 100
                    .TextAlign = fmTextAlignLeft
                    .Caption = g_strTxt(153) & ": " & Format(st, "0.00") & " " & g_strTxt(151) 'Stake: xx EUR
                End With
                
            'Pay-out
                Set objLabel7 = frmBettingAnalysis.Controls.Add("Forms.Label.1", , True)
                With objLabel7
                    .Font.name = "Tahoma"
                    .Font.Size = 10
                    .left = 200
                    .top = 105 + k * 12
                    .Width = 150
                    .TextAlign = fmTextAlignLeft
                    .Caption = "  " & g_strTxt(154) & ": " & Format(payCash, "0.00") & " " & g_strTxt(151) 'Pay-out: xx EUR
                    .BackColor = payColor
                End With
                
                k = k + 2
            
        Next i

        With frmBettingAnalysis
            .Caption = m_strRaceName & " " & m_strRaceYear & " | " & m_strRaceTrack & ", " & m_strRaceLocation _
                        & " (" & m_strRaceCountry & ")"
            .Width = 400 'width of the pop-up

            If g_colBetSlips.Count <= 5 Then
                .Height = 105 + k * 12 + 30 'height of the pop-up depending of the number of bets placed
            Else
                .Height = 440 'fix height if more than 5 bets are placed
            End If
            .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
            .ScrollHeight = 105 + k * 12 'height of the vertical scrolling
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
    strWarningMessage = strWarningMessage & g_strTxt(30) & vbNewLine & g_strTxt(31)

    'Pop-up
        'Set the button mode
        g_strMsgButtons = "OK"
        'Assign the text for the pop-up
        g_strMsgCaption = g_c_TOOL
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
'            g_wksMovie.Delete
        Application.DisplayAlerts = True 'Turn on application warnings
    On Error GoTo 0
End Sub

'Information for UserForm "frmStart"
Private Sub UserFormSTART()
    With frmStart
        .Caption = g_c_TOOL
        .lblS1.Caption = m_strRaceName & " " & m_strRaceYear 'race name and year
        .lblS2.Caption = m_strRaceType & " " & g_strTxt(1) & " " & m_intTrackLength & " " & g_strTxt(3) ' race type and distance
        .lblS3.Caption = m_strRaceTrack & " " & g_strTxt(10) & " " & m_strRaceLocation & " (" & m_strRaceCountry & ")" 'race course, loaction and country
        .lblS6.Caption = m_intHorsesStarting & " " & g_strTxt(100) 'number of horses starting
        If m_strRealRace = "Y" Then
            .lblS10.Caption = UCase(g_strTxt(108)) 'REAL RACE
        Else
            .lblS10.Caption = UCase(g_strTxt(109)) 'FICTITIOUS RACE
        End If
        .lblS8.Caption = g_strTxt(104) 'caption of the betting section
        .lblFocus.Caption = g_strTxt(105) 'label "Focused horse"
        .cmdS1.Caption = g_strTxt(101) 'button "Add bet slip"
        .cmdS2.Caption = g_strTxt(102) 'button "Start race"
        With .lblS4 'track surface
            .Caption = m_strTrackSurface 'text with the surface
            .BorderStyle = fmBorderStyleSingle 'draw a border around
            .BackColor = m_lngTrackColour 'set the color according to the track colour
        End With
        .cmdS6.Caption = g_strTxt(107) 'button "Show horse numbers and odds"
        If g_blnBettingMode Then 'height of the pop-up
            .Height = 290 'if the betting mode is enabled
        Else
            .Height = 180 'if the betting mode is disabled
        End If
        .Show (vbModal) 'modal
    End With
End Sub

'Name of the gambler for placing a bet
Public Sub Gambler()
    
    'Set the button mode
    g_strMsgButtons = "CancelOK"
    'Assign the text for the pop-up
    g_strMsgText = g_strTxt(111)
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
    Dim objLabel1 As MSForms.Label 'horse nr
    Dim objLabel2 As MSForms.Label 'horse name
    Dim objLabel3 As MSForms.Label 'estimated horse condition (+/-50)
    Dim objLabel4 As MSForms.Label 'real condition (for development purposes only)
    
    Dim min As Integer, max As Integer
    Dim i As Integer, j As Integer, k As Integer
    
    With frmOdds
        .Caption = m_strRaceName & " " & m_strRaceYear
        .Height = 60
        .lblO0.Caption = g_strTxt(160)
        .lblO1.Caption = g_strTxt(161)
        .lblO2.Caption = g_strTxt(162)
        .lblO3.Caption = g_strTxt(163)
    End With
    
    k = 1

    For i = 1 To UBound(g_arr_varHorses)
        If min = 0 Or g_arr_varHorses(i, 17) < min Then min = g_arr_varHorses(i, 17)
        If g_arr_varHorses(i, 17) > max Then max = g_arr_varHorses(i, 17)
    Next i
    
    For i = min To max
        For j = 1 To UBound(g_arr_varHorses)
            If g_arr_varHorses(j, 17) = i Then
            
                Set objLabel1 = frmOdds.Controls.Add("Forms.Label.1", , True)
                With objLabel1 'Nr and name of the horse
                    .Font.name = "Tahoma"
                    .Font.Size = 12
                    .left = 12
                    .top = 12 + k * 18
                    .Width = 200
                    .TextAlign = fmTextAlignLeft
                    .Caption = "#" & g_arr_varHorses(j, 11) & vbTab & g_arr_varHorses(j, 1)
                    If g_arr_varHorses(j, 0) <> "START" Then .Font.Strikethrough = True
                End With
                
                'Adjust UserForm height
                frmOdds.Height = frmOdds.Height + objLabel1.Height

                Set objLabel2 = frmOdds.Controls.Add("Forms.Label.1", , True)
                With objLabel2 'odd
                    .Font.name = "Tahoma"
                    .Font.Size = 12
                    .left = 220
                    .top = 12 + k * 18
                    .Width = 62
                    .TextAlign = fmTextAlignRight
                    .Caption = g_arr_varHorses(j, 17) & ":10"
                    If g_arr_varHorses(j, 0) <> "START" Then .Font.Strikethrough = True
                End With
                
                If g_arr_varHorses(j, 0) = "START" Then
                    Set objLabel3 = frmOdds.Controls.Add("Forms.Label.1", , True)
                    With objLabel3 '~condition
                        .Font.name = "Tahoma"
                        .Font.Size = 12
                        .left = 290
                        .top = 13 + k * 18
                        .Height = 15
                        .Width = 100 + (((g_arr_varHorses(j, 6) * 1000000) - 1499900) / 2) + g_arr_varHorses(j, 18)
                        .TextAlign = fmTextAlignLeft
                        .BackColor = 6740479
''For development purposes only
'                        .Caption = "T: " & g_arr_varHorses(j, 6) * 1000000 & " ... " & g_arr_varHorses(j, 18)
                    End With
                End If

''For development purposes only
'                Set objLabel4 = frmOdds.Controls.Add("Forms.Label.1", , True)
'                With objLabel4 'real condition
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

    frmOdds.Show (vbModal) 'modal
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
        .lblR1 = UCase(m_strRaceLocation & " (" & m_strRaceCountry & ")")
        .lblR2 = UCase(m_strRaceName)
        .lblR3 = m_intHorsesStarting & " " & UCase(g_strTxt(100))
        .lblR4 = UCase(g_colBetSlips(id).BType)
        .lblR5 = UCase(bet)
        .lblR6 = UCase(g_strTxt(152) & " " & Format(g_colBetSlips(id).Stake, "0.00") & " " & g_strTxt(151))
        .lblR7 = g_colBetSlips(id).id
        .Show (vbModal) 'modal
    End With
End Sub


'CALLBACKS (Excel ribbon events)
'-------------------------------

'Callback for customUI.onLoad
Public Sub AI_GaloppSimAddinInitialize(ribbon As IRibbonUI)
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
Public Sub AI_GetLabel(control As IRibbonControl, ByRef returnedVal)
    Select Case control.id
    
        Case "group01GALOPPSIM" 'group "Settings"
            returnedVal = g_strTxt(360)
            Case "menu01GALOPPSIM" 'menu "Language"
                returnedVal = g_strTxt(229)
                Case "btn01aGALOPPSIM" 'button "deutsch"
                    returnedVal = g_strTxt(230)
                Case "btn01bGALOPPSIM" 'button "english"
                    returnedVal = g_strTxt(231)
            Case "btn10GALOPPSIM" 'button "Race options"
                returnedVal = g_strTxt(252)
            Case "btn20GALOPPSIM" 'button "Excel options"
                returnedVal = g_strTxt(253)
                
        Case "group02GALOPPSIM" 'group "Race"
            returnedVal = g_strTxt(361)
            Case "combo01InstalledRaces" 'combobox "Race selection"
                returnedVal = g_strTxt(288)
            Case "btn30GALOPPSIM" 'button "Start race"
                returnedVal = g_strTxt(278)
            Case "btn31GALOPPSIM" 'button "Photo of the finish"
                returnedVal = g_strTxt(280)
            Case "btn32GALOPPSIM" 'button "Ranking list"
                returnedVal = g_strTxt(279)
            Case "btn33GALOPPSIM" 'button "Photo of the winner"
                returnedVal = g_strTxt(286)
            Case "btn34GALOPPSIM" 'button "Betting analysis"
                returnedVal = g_strTxt(281)
            
        Case "btn40GALOPPSIM" 'button "Info"
            returnedVal = g_strTxt(282)
        Case "btn50GALOPPSIM" 'button "Warning"
            returnedVal = g_strTxt(283)
        Case "btn60GALOPPSIM" 'button "Movie"
            returnedVal = g_strTxt(284)
        Case "btn70GALOPPSIM" 'button "Close"
            returnedVal = g_strTxt(285)
    End Select
End Sub

'Callbacks for Tooltips
Public Sub AI_GetScreentip(control As IRibbonControl, ByRef screentip)
   
'ToDo: v148.60

End Sub

Public Sub AI_GetSupertip(control As IRibbonControl, ByRef supertip)
   
'ToDo: v148.60

End Sub

'Ribbon button "Race options"
Public Sub AI_OptionsRace(control As IRibbonControl)
    frmOptionsRace.Show (vbModal) 'display UserForm (modal)
End Sub

'Ribbon button "Excel options"
Public Sub AI_OptionsExcel(control As IRibbonControl)
    'REPLACE in v148.60 ToDo!!!!!
        'Set the button mode
        g_strMsgButtons = "OK"
        'Assign the text for the pop-up
        g_strMsgCaption = g_c_TOOL
        g_strMsgText = g_strTxt(50)
        'Display the pop-up
        frmMsg_MultiPurpose.Show (vbModal) 'modal
'    frmOptionsExcel.Show (vbModal) 'display UserForm (modal)
End Sub

'Startbutton im Menüband
Public Sub AI_StartRace(control As IRibbonControl)
    'Leaving the current race?
    If g_blnRaceStarted Then
        Dim strNextRace As String
        With ThisWorkbook.Worksheets(g_strRaceSelected)
            strNextRace = .Cells(4, 2).Value & " " & _
                .Cells(5, 2).Value & " (" & _
                .Cells(12, 2).Value & "m) - " & _
                .Cells(6, 2).Value
        End With
        
'        'A MessageBox cannot handle unicode, so for example Cyrillic characters are displayed as question marks
'        If MsgBox((g_strTxt(11) & " " & g_wksRace.OLEObjects("CBraces").Object.text), _
'            vbOKCancel, g_c_TOOL) = vbCancel Then Exit Sub

        'Pop-up
            'Set the button mode
            g_strMsgButtons = "CancelOK"
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_TOOL
            g_strMsgText = g_strTxt(11) & " " & strNextRace
            'Display the pop-up
            frmMsg_MultiPurpose.Show (vbModal) 'modal
            'Evaluate the return value
            If g_strButtonPressed = "CANCEL" Then Exit Sub
            
            Call ShowNewRaceScreen
    End If

    Call NewRace
End Sub

'Ergebnis-Button im Menüband
Public Sub AI_Results(control As IRibbonControl)
    Call show_ergebnis
End Sub

'Ribbon button "Winner"
Public Sub AI_Winner(control As IRibbonControl)
    Call show_winnerPhoto
End Sub

Private Sub show_winnerPhoto()
    If g_blnRaceStarted Then
        g_wksRace.Activate
        Call basAuxiliary.Scroll(m_intTrackLength + m_intLeftColumns + 9 + 160, m_intTopRows + (g_intHorsesEnrolled * 2 + 9))
        If g_strPlayMode = "RS" Then frmRS_navigation.Show (vbModeless) 'modeless
    Else
        Call AssignDataSheets
        Call GetTextComponents 'Texte einlesen
        'Pop-up
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_TOOL
            g_strMsgText = g_strTxt(156)
            'Display the pop-up
            frmMsg_Attention.Show (vbModal) 'modal
    End If
End Sub

Private Sub show_ergebnis()
    If g_blnRaceStarted Then
        g_wksRace.Activate
        Call basAuxiliary.Scroll(m_intTrackLength, m_intTopRows + (g_intHorsesEnrolled * 2 + 9))
        If g_strPlayMode = "RS" Then frmRS_navigation.Show (vbModeless) 'modeless
    Else
        Call AssignDataSheets
        Call GetTextComponents 'Texte einlesen
        'Pop-up
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_TOOL
            g_strMsgText = g_strTxt(156)
            'Display the pop-up
            frmMsg_Attention.Show (vbModal) 'modal
    End If
End Sub

'Zielfoto-Button im Menüband
Public Sub AI_FinishPhoto(control As IRibbonControl)
    Call show_foto
End Sub

Private Sub show_foto()
    If g_blnRaceStarted Then
        g_wksRace.Activate
        'Text anpassen
            g_wksRace.Cells(4 + m_intTopRows, m_intTrackLength + m_intLeftColumns + 9).Value = g_strTxt(28) '("Zielfoto")
        Call basAuxiliary.Scroll(m_intTrackLength, m_intTopRows + 1)
        Call ShowFinishPhoto
        If g_strPlayMode = "RS" Then frmRS_navigation.Show (vbModeless) 'modeless
    Else
        Call AssignDataSheets
        Call GetTextComponents 'Texte einlesen
        'Pop-up
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_TOOL
            g_strMsgText = g_strTxt(156)
            'Display the pop-up
            frmMsg_Attention.Show (vbModal) 'modal
    End If
End Sub

'Wett-Button im Menüband
Public Sub AI_Betting(control As IRibbonControl)
    Call show_wetten
End Sub

Private Sub show_wetten()
    If g_blnRaceStarted And g_blnBetsPlaced Then
        Call UserFormAnalyseBetSlips
    Else
        Call AssignDataSheets
        Call GetTextComponents 'Texte einlesen
        'Pop-up
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_TOOL
            g_strMsgText = g_strTxt(155)
            'Display the pop-up
            frmMsg_Attention.Show (vbModal) 'modal
    End If
End Sub

'Info-Button im Menüband
Public Sub AI_Info(control As IRibbonControl)
    Call show_info
End Sub

Private Sub show_info()
    With frmInfo
        .Caption = g_c_TOOL
        .lbl_info01.Caption = g_strTxtInfo(1) & vbNewLine _
                            & g_strTxtInfo(2) & vbNewLine _
                            & g_strTxtInfo(3)
        'MultiPage captions
        For i = 0 To 3
            .multiPage_info.Pages(i).Caption = g_strTxtInfo(10 + i)
        Next i
            .multiPage_info.Value = 0 'set the focus on the first page
        'Page "GaloppSim info"
            With .lbl_info_galoppsim01
                .Caption = g_strTxtInfo(100) & vbNewLine & vbNewLine _
                            & g_strTxtInfo(102) & vbNewLine & vbNewLine _
                            & g_strTxtInfo(104) & vbNewLine & vbNewLine _
                            & g_strTxtInfo(106) & vbNewLine & vbNewLine _
                            & g_strTxtInfo(108) & vbNewLine & vbNewLine _
                            & g_strTxtInfo(110) & vbNewLine & vbNewLine _
                            & g_strTxtInfo(112) & vbNewLine & vbNewLine _
                            & g_strTxtInfo(114) & vbNewLine & vbNewLine _
                            & g_strTxtInfo(116) & vbNewLine & vbNewLine _
                            & g_strTxtInfo(118) & vbNewLine & vbNewLine
                .Width = 460 'label width fix
                .AutoSize = True 'label height depending on the content
            End With
            With .multiPage_info.Pages(0)
                .ScrollBars = fmScrollBarsVertical 'vertical scrollbar
                .ScrollHeight = frmInfo.lbl_info_galoppsim01.Height 'height of the vertical scrolling
                .KeepScrollBarsVisible = fmScrollBarsNone 'show scrollbars only when needed
            End With
        'Page "Source code"
            With .lbl_info_sourcecode01
                .Caption = g_strTxtInfo(30)
                .Font.Size = 12
                .WordWrap = True
                .AutoSize = True 'label height depending on the content
            End With
            .btn_info_sourcecode01.ControlTipText = g_strTxtInfo(31)
        'Page "Contact"
            With .lbl_info_contact01
                .Caption = g_strTxtInfo(50) & ":" & vbNewLine & vbNewLine _
                            & g_strTxtInfo(51) & vbNewLine _
                            & g_strTxtInfo(52) & vbNewLine & vbNewLine _
                            & g_strTxtInfo(53)
                .Font.Size = 12
                .WordWrap = True
            End With
            With .btn_info_contact01
                .Caption = g_strTxtInfo(54)
                .ControlTipText = g_strTxtInfo(55)
                .Font.Size = 12
                .WordWrap = True
            End With
        'Page "Donation"
            With .lbl_info_donation01
                .Font.Size = 12
                .Caption = g_strTxtInfo(40) & vbNewLine & vbNewLine _
                            & g_strTxtInfo(41)
                .AutoSize = True 'label height depending on the content
            End With
            .btn_info_donation01.ControlTipText = g_strTxtInfo(42)
            .btn_info_donation02.ControlTipText = g_strTxtInfo(43)
        .Show (vbModal) 'modal
    End With
End Sub

'Warnhinweis-Button im Menüband
Public Sub AI_Warning(control As IRibbonControl)
    Call show_warning
End Sub

Private Sub show_warning()
    Call AssignDataSheets
    Call GetTextComponents
    Call ShowWarning
End Sub

'Movie button (Ribbon)
Public Sub AI_Movie2017(control As IRibbonControl)
    Call AssignDataSheets
    Call movie2017
End Sub

Private Sub movie2017()
    'Play the movie
    Call basMovie2017.PlayMovie2017(g_strTxt(), ActiveSheet.name)
End Sub

'Ende-Button im Menüband
Public Sub AI_Close(control As IRibbonControl) 'ToDo... in v148.60 komplett durchdenken und überarbeiten!!
    If g_blnRaceStarted Then g_blnRaceStarted = False
    'Reset Excel options
        Call ResetExcelOptions 'ToDo ist da nötig??
        Application.ScreenUpdating = True
    'Blatt löschen
        Call AI_DeleteWorksheets
    Unload frmBettingAnalysis 'ToDo sauber machen, evtl. in eigener prozedur
End Sub

'Language button "DE"
Public Sub AI_LanguageDE(control As IRibbonControl)
    g_strLanguage = "DE"
    Call ChangeLanguage
End Sub

'Language button "EN"
Public Sub AI_LanguageEN(control As IRibbonControl)
    g_strLanguage = "EN"
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
    Select Case name
        'Button labels
        Case "startrace"
            g_wksRace.OLEObjects(name).Object.Caption = g_strTxt(278)
        Case "fotofinish"
            g_wksRace.OLEObjects(name).Object.Caption = g_strTxt(280)
        Case "results"
            g_wksRace.OLEObjects(name).Object.Caption = g_strTxt(279)
        Case "winner"
            g_wksRace.OLEObjects(name).Object.Caption = g_strTxt(286)
        Case "bets"
            g_wksRace.OLEObjects(name).Object.Caption = g_strTxt(281)
        Case "raceoptions"
            g_wksRace.OLEObjects(name).Object.Caption = g_strTxt(252)
        Case "exceloptions"
            g_wksRace.OLEObjects(name).Object.Caption = g_strTxt(253)
        Case "language"
            g_wksRace.OLEObjects(name).Object.Caption = g_strTxt(229)
        Case "info"
            g_wksRace.OLEObjects(name).Object.Caption = g_strTxt(282)
        Case "warning"
            g_wksRace.OLEObjects(name).Object.Caption = g_strTxt(283)
        Case "movie2017"
            g_wksRace.OLEObjects(name).Object.Caption = g_strTxt(284)
    End Select
End Sub

'Anzahl der installierten Rennen auslesen
Public Sub AI_InstalledRaces_getItemCount(control As IRibbonControl, ByRef returnedVal)
    
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
Public Sub AI_InstalledRaces_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)

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

'Place the UserForm in the center of the Window
    'frmMe as UserForm geht nicht: http://www.office-loesung.de/ftopic516038_0_0_asc.php
Public Sub PlaceUserFormInCenter(frmMe As Object)
    With frmMe
        .StartUpPosition = 0
        .top = ActiveWindow.top + ((ActiveWindow.Height - frmMe.Height) / 2)
        .left = ActiveWindow.left + ((ActiveWindow.Width - frmMe.Width) / 2)
    End With
End Sub

'Ausgewähltes Rennen setzen
Public Sub AI_InstalledRaces_Click(control As IRibbonControl, id As String, index As Integer)

    g_strRaceSelected = m_colRacesInstalled(index + 1)

End Sub


'ToDo: Alternative zu Workbook_BeforeClose... Beschreiben im Buch welches Ereignis zuerst ausgeführt wird...
'Execute this procedure when the workbook is being closed
Public Sub Auto_Close()
    If g_strPlayMode = "RS" Then
        'Reset Excel options
            Call basMainCode.ResetExcelOptions
            Application.ScreenUpdating = True
        'Do not save the workbook
        'https://support.microsoft.com/de-de/help/213428/how-to-suppress-save-changes-prompt-when-you-close-a-workbook-in-excel
        ThisWorkbook.Saved = True
    End If
End Sub
