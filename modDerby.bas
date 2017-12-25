Attribute VB_Name = "modDerby"
Option Explicit
Option Private Module

'Horse racing simulator (Galopp-Simulator) - version 148.30 - December 2017
'Marco Krapf - excel@marco-krapf.de - https://marco-krapf.de/excel/
'GNU General Public License v3.0

'Workbookweit gültige Variablen und Konstanten
Public Const tool As String = "GaloppSim (v148.30)"

'Workbookweit gültige Variablen für Menüband-Elemente
Public GaloppSimRibbon As IRibbonUI 'Custom ribbon
Public language As String 'Language
Public PlayMode As String ' "AI" = AddIn (xlam) "RS" = Run simple Workbook (xlsb)
Public collRSbuttons As Collection 'Buttons in RS edition
Public collRScheckboxes As Collection 'CheckBoxes in RS edition
Public comboRace As OLEObject 'ComboBox with races in RS edition
Public comboBet As OLEObject 'ComboBox with bet slips in RS edition
Public raceID As String 'Unique race ID
Public rennen As Boolean 'Kennzeichen ob ein Rennen gestartet wurde
Public maxi As Boolean 'Kennzeichen ob Fenster auf Bildschirmgröße maximiert wird
Public gitter As Boolean 'Kennzeichen ob Gitternetzlinien auf dem Tabellenblatt angezeigt werden
Public headings As Boolean 'Kennzeichen ob die Zeilen/Spaltennummern auf dem Tabellenblatt angezeigt werden
Public zoom As Boolean 'Kennzeichen ob Rennbahn gezoomt werden soll wenn das Fenster ausreichend groß ist
Public namen As Boolean 'Kennzeichen ob die Namen der Pferde am linken Rand fix angezeigt werden
Public namen2 As Boolean 'Kennzeichen ob die Namen der Pferde im Ziel angezeigt werden
Public pferdefarbe As Boolean 'Kennzeichen ob die Farbe des Pferds am linken Rand angezeigt werden soll
Public hufspur As Boolean 'Kennzeichen ob Hufspuren angezeigt werden (verlangsamt das Rendering)
Public taktik As Boolean 'Kennzeichen ob die Pferde pro Rennphase unterschiedlich schnell sind
Public spannung As Boolean 'Kennzeichen ob die Ergebnisse auf der Tafel mit kleiner Pause angezeigt werden sollen
Public markFav As Boolean 'Kennzeichen ob der Favorit markiert werden soll
Public colResult As Boolean 'Show/hide horse colours in results
Public betting As Boolean 'Betting or not
Public Auswahl As String 'Ausgewähltes Rennen
Public arrayPferde() As Variant 'Alle Infos über jedes Pferd
Public collSlips As Collection 'All bet slips
Public txt(1 To 500) As String 'Array für Textbausteine
Public wks As Worksheet 'Tabellenblatt für das Rennen

'Modulweit gültige Variablen
Dim wksData As Worksheet 'Tabellenblatt mit den Renndaten
Dim wksTxt As Worksheet 'Tabellenblatt mit Texten
Dim wksPic As Worksheet 'Worksheet with picture data
Dim wksAdv As Worksheet 'Worksheet with advertising data (gsadv file format)
Dim wkscheck As Worksheet 'Variable zum Prüfen ob das Tabellenblatt GALOPPSIM da ist
Dim arrayFotofinish() As Variant 'Position jedes Pferds wenn das erste im Ziel ankommt
Dim arrayBerechnung() As Variant 'Berechnung der Platzierung beim (gleichzeitigen) Zieleinlauf
Dim arrayErgebnisse() As Variant 'Ergebnisliste
Dim collRennen As Collection 'Installierte Rennen
Dim collBet As Collection 'Collection with bet slips
Dim arrayAdv() As Variant 'Advertisement
Dim test As Boolean 'Kennzeichen ob das Programm im Testmodus während der Entwicklung läuft
Dim infoblatt As Boolean 'Kennzeichen ob die Info angezeigt wurde
Dim race As String 'Name des Rennens
Dim rennort As String 'Ort der Rennbahn
Dim rennbahn As String 'Name der Rennbahn
Dim jahr As String 'Jahr des Rennens
Dim gelaeuf As String 'Art der Rennbahn (Grasbahn, Sandbahn, Schnee)
Dim trackCol As Long 'Colour of the track surface
Dim raceType As String 'Flachrennen oder Hindernisrennen
Dim raceTypeID As String 'Race type (F = flat, S = steeplechase...)
Dim randomLane As String 'lanes fix or random (F/R)
Dim randomCol As String 'horse colours fix or random (F/R)
Dim randomOdd As String 'odds fix or random (F/R)
Dim advYN As String 'Advertising (Y/N)
Dim topRows As Integer 'Number of rows at the top (used for RS mode)
Dim vorlauf As Integer 'Anzahl Spalten vor den Startboxen
Dim meter As Integer 'Länge der Rennbahn in Metern
Dim spur As Integer 'Breite der Spur (in Abhängigkeit der Fenstergröße)
Dim spur2 As Double 'Länge der Spur (in Abhängigkeit der Fenstergröße)
Dim liste As Double 'Breite Zielfoto/Ergebnistafel (in Abhängigkeit der Fenstergröße)
Dim scrolling As Integer 'Scrolling (in Abhängigkeit der Fenstergröße)
Dim schrift As Integer 'Schriftgröße (in Abhängigkeit der Fenstergröße)
Dim speedF1 As Long, speedF2 As Long 'Variablen für die Geschwindigkeiten der Form der Pferde
Dim speedK1 As Long, speedK2 As Long 'Variablen für den kursfristigen Faktor pro Schleifendurchlauf
Dim speedP1 As Long, speedP2 As Long, speedP3 As Long 'Variablen für die Geschwindigkeiten der Renntaktik
Dim startplaetze As Integer 'Anzahl der Pferde, die gemeldet sind
Dim starter As Integer 'Anzahl der Pferde, die starten
Dim favorit1 As Integer, favorit2 As Integer, favorit3 As Integer 'Favoriten für das Rennem
Dim favsum1 As Double, favsum2 As Double, favsum3 As Double 'Für Berechnung der Favoriten
Dim einlauf As Integer 'Zählvariable, wie viele Pferde in einem Schleifendurchlauf im Ziel ankommen
Dim platz As Integer 'Platzierung im Rennen
Dim sieger As String 'Name(n) des Pferds bzw. der Pferde auf Platz 1
Dim sieg As Boolean, fotofin As Boolean 'Variablen zur Ermittlung, ob es ein Fotofinish gibt
Dim i As Integer, j As Integer, k As Integer, z As Long

'Starting procedure (RS edition)
Public Sub derby_RS()
    ThisWorkbook.Worksheets("RS edition").Activate
    topRows = 8
    
    Call Grunddaten2 'Assign worksheets "Txt", "Adv" and "Pic"
    Call Texte 'Get texts
    Call Grunddaten1 'Create worksheet "GALOPPSIM"
    Call RS_StartScreen 'Startscreen with Loppsi and navigation panel
    Call RS_headings 'Menu headings
    Call RS_controls 'Add controls
    Call RS_raceSelect 'Add dropdown for race selection
End Sub

'Starting procedure
Private Sub derby()
    'Fehlerbehandlung
    On Error GoTo FEHLER
        Cells.Clear 'Clear the whole worksheet
        
        If PlayMode = "AI" Then
            Call Grunddaten2 'Tabellenblätter "Txt", "Adv" and "Pic"
            Call Texte 'Texte einlesen
        End If
    
        If AuswahlCheck(Auswahl) = False Then
            MsgBox txt(98), vbExclamation, tool
            Exit Sub
        End If
        
        Set collSlips = Nothing
        Set collSlips = New Collection
    
        Call Renndaten 'Tabellenblatt mit ausgewähltem Rennen
        Call Grunddaten4 'Grundsätzliche Daten auslesen bzw. festlegen
        Call Pferdedaten 'Daten über das Rennen einlesen
        Call UserFormSTART 'Show Start-UserForm
        
    If rennen Then
        If PlayMode = "AI" Then
            Call Grunddaten1 'Tabellenblatt "GALOPPSIM"
        End If
        
        If PlayMode = "RS" Then 'Hide navigation area when running RS mode
            Call RS_HideNavi
        End If

        Call Rennstrecke 'Geläuf zeichnen
        Call Startfeld 'Pferde generieren
        Call Begruessung 'Popup zu Rennbeginn
        Call Startaufstellung 'Pferde in Boxen stellen
        Call Vorstellung 'Pferde vorstellen
        Call Rennstart 'Rennstart
        Call Ergebnisse 'Ergebnistafel
        Call Siegerpferd 'Grafik
        If betting Then Call AnalyseBetSlips 'Analyse bet slips
        
        If PlayMode = "RS" Then 'Show navigation area when running RS mode
            'Activate buttons
            With wks
                .OLEObjects("results").Object.Enabled = True
                .OLEObjects("fotofinish").Object.Enabled = True
                If betting Then
                    .OLEObjects("bets").Object.Enabled = True
                End If
            End With
            Call RS_ShowNavi
        End If
    End If

    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Standardparameter setzen für Tests während der Entwicklung
Private Sub Standardparameter()
    language = "DE"
    taktik = True
    namen = True
    namen2 = True
    pferdefarbe = True
    hufspur = False
    spannung = False
    maxi = True
    gitter = False
    headings = False
    zoom = True
    betting = True
    markFav = True
    colResult = True
End Sub

'AuswahlCheck
Private Function AuswahlCheck(a As String) As Boolean
    If a = "" Then
        AuswahlCheck = False
    Else
        AuswahlCheck = True
    End If
End Function

Public Sub RS_raceSelect()
    'Add combobox to the navigation panel
    Call RS_addComboboxRaces("CBraces", 18, 18, 251, 24)
End Sub

'Add controls in RS mode
Public Sub RS_controls()
    'Error handling
    On Error GoTo FEHLER
    
    'Add buttons to the navigation panel
        Set collRSbuttons = New Collection
        
        Call RS_addButton("derby", 18, 43, 81, 49, txt(184)) 'txt(...) is the initial caption
        Call RS_addButton("results", 102, 43, 81, 49, txt(185))
        Call RS_addButton("fotofinish", 186, 43, 81, 49, txt(186))
        Call RS_addButton("bets", 289, 43, 135, 49, txt(187))
        Call RS_addButton("info", 1001, 22, 94, 22, txt(188))
        Call RS_addButton("warning", 1001, 46, 94, 22, txt(189))
        Call RS_addButton("movie1", 1001, 70, 94, 22, txt(220))
        
        'Language
        Call RS_addButton("DE", 18, 100, 81, 22, txt(230))
        Call RS_addButton("EN", 102, 100, 81, 22, txt(231))
'        Call RS_addButton("RU", 186, 100, 81, 22, txt(232))

        'name, left, top, width, height, caption, [optional state]
        
        'Racing Options
        Call RS_addButton("tac", 448, 22, 124, 22, txt(194))
        Call RS_addButton("huf", 448, 46, 124, 22, txt(195))
        Call RS_addButton("nameS", 448, 70, 124, 22, txt(196))
        Call RS_addButton("colS", 574, 22, 124, 22, txt(197))
        Call RS_addButton("mFav", 574, 46, 124, 22, txt(198))
        Call RS_addButton("nameF", 574, 70, 124, 22, txt(199))
        Call RS_addButton("colR", 700, 22, 124, 22, txt(210))
        Call RS_addButton("vzRes", 700, 46, 124, 22, txt(201))
        Call RS_addButton("zoom", 700, 70, 124, 22, txt(202))
        'Excel options
        Call RS_addButton("maxi", 847, 22, 136, 22, txt(190))
        Call RS_addButton("RCHide", 847, 46, 136, 22, txt(214))
        Call RS_addButton("GridHide", 847, 70, 136, 22, txt(215))

        Call RS_buttons_inactivate 'inactivate some buttons
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

Private Sub RS_buttons_inactivate()
    'Set buttons inactive
        With wks
            .OLEObjects("results").Object.Enabled = False
            .OLEObjects("fotofinish").Object.Enabled = False
            .OLEObjects("bets").Object.Enabled = False
        End With
End Sub

'Click on RS button
Public Sub RS_execute_Click(name As String)
    
    Dim txTr As Integer, txFa As Integer
    
    On Error GoTo NORACE
    
    Select Case name

        Case "derby"
            If comboRace.Object.Value = "" Then
                MsgBox txt(98), vbExclamation, tool
                Exit Sub
            Else
                'Leaving the current race?
                If rennen Then
                    If MsgBox((txt(11) & " " & wks.OLEObjects("CBraces").Object.Text), _
                        vbOKCancel, tool) = vbCancel Then Exit Sub
                End If
                
                Auswahl = comboRace.Object.Value 'get race from dropdown
            End If
            
            'Clear start screen colors
            wks.Range(Columns(1), Columns(80)).Interior.Color = xlNone
            
            Call RS_buttons_inactivate 'inactivate some buttons
            Call derby
        
        Case "results"
                Call ergebnis_show
        
        Case "fotofinish"
                Call foto_show
        
        Case "bets"
                Call wetten_show
            
        Case "info"
            Call info_show
            
        Case "warning"
            Call warning_show
            
        Case "movie1"
            Call movie1
            
        Case "DE"
            Call languageDE
        Case "EN"
            Call languageEN
        Case "RU"
            Call languageRU
        
    'Options
        Case "bet"
            '... = OLEObject, state true/false, caption if false, caption if true
            betting = RS_ButtonState(wks.OLEObjects(name), betting, 203, 193)
            
        Case "tac"
            taktik = RS_ButtonState(wks.OLEObjects(name), taktik, 204, 194)
            
        Case "huf"
            hufspur = RS_ButtonState(wks.OLEObjects(name), hufspur, 205, 195)

        Case "nameS"
            namen = RS_ButtonState(wks.OLEObjects(name), namen, 206, 196)
            
        Case "colS"
            pferdefarbe = RS_ButtonState(wks.OLEObjects(name), pferdefarbe, 207, 197)
            
        Case "mFav"
            markFav = RS_ButtonState(wks.OLEObjects(name), markFav, 208, 198)
            
        Case "nameF"
            namen2 = RS_ButtonState(wks.OLEObjects(name), namen2, 209, 199)
            
        Case "colR"
            colResult = RS_ButtonState(wks.OLEObjects(name), colResult, 210, 200)
            
        Case "vzRes"
            spannung = RS_ButtonState(wks.OLEObjects(name), spannung, 211, 201)
            
        Case "zoom"
            zoom = RS_ButtonState(wks.OLEObjects(name), zoom, 212, 202)
            
        Case "maxi"
            maxi = RS_ButtonState(wks.OLEObjects(name), maxi, 213, 190)
            
        Case "RCHide"
            headings = RS_ButtonState(wks.OLEObjects(name), headings, 214, 191)
            
        Case "GridHide"
            gitter = RS_ButtonState(wks.OLEObjects(name), gitter, 215, 192)
            
    End Select
    
    Exit Sub

NORACE:
    MsgBox txt(98), vbExclamation, tool
End Sub

'Add new RS frame
Private Sub RS_addFrame(n As String, l As Integer, t As Integer, _
                        w As Integer, h As Integer, c As String)
    Dim RSframe As OLEObject
    
    Set RSframe = wks.OLEObjects.Add(classtype:="Forms.Frame.1", _
    left:=l, top:=t, width:=w, height:=h)
    
    With RSframe
        .name = n
        .Object.BackColor = 255 '&HFFFFFF
        .Object.Caption = c
        .Placement = xlFreeFloating
        .Visible = True
    End With
 
End Sub

'Add new RS button
Private Sub RS_addButton(n As String, l As Integer, t As Integer, w As Integer, h As Integer, _
                        c As String)
    Dim RSbtn As OLEObject
    Dim RSbutton As clsRSbutton
    
    Set RSbtn = wks.OLEObjects.Add(classtype:="Forms.CommandButton.1", _
    left:=l, top:=t, width:=w, height:=h)
    
    With RSbtn
        .name = n
        .Object.Caption = c
        .Placement = xlFreeFloating
        .Object.TakeFocusOnClick = False
        .Visible = True
    End With
    
    Set RSbutton = New clsRSbutton
    Set RSbutton.RSButtonObject = RSbtn.Object
    RSbutton.RSbtnID = n

    collRSbuttons.Add RSbutton
 
End Sub

'Set/change RS button state and caption
Private Function RS_ButtonState(RSbtn As OLEObject, state As Boolean, txFa As Integer, txTr As Integer)
    Dim textC As String
    If state = True Then
        textC = txt(txFa)
        state = False
    Else
        textC = txt(txTr)
        state = True
    End If
    RSbtn.Object.Caption = textC
    RS_ButtonState = state
End Function

'Add new RS combobox with races
Private Sub RS_addComboboxRaces(n As String, l As Integer, t As Integer, w As Integer, h As Integer)

    Dim wksR As Worksheet
    Dim strR As Variant
    
    Set collRennen = Nothing
    Set collRennen = New Collection

    Set comboRace = wks.OLEObjects.Add(classtype:="Forms.ComboBox.1", _
        left:=l, top:=t, width:=w, height:=h)
    
    With comboRace
        .name = n
        .Placement = xlFreeFloating
        .Object.ColumnCount = 2 'column 0: wks-name // column 1: visible name
        .Object.ColumnWidths = "0 Pt" 'width of the column with the race name --> hidden
        .Visible = True
    End With
    
    'Populate Dropdown with installed races
    For Each wksR In ThisWorkbook.Worksheets
        If left(wksR.name, 5) = "race_" Then
            collRennen.Add wksR.name
            With comboRace.Object
            .AddItem
            .List(.ListCount - 1, 0) = wksR.name
            .List(.ListCount - 1, 1) = wksR.Cells(2, 2).Value & " " & _
                                    wksR.Cells(3, 2).Value & " (" & _
                                    wksR.Cells(7, 2).Value & "m)"
            End With
        End If
    Next wksR
 
End Sub

'Add new RS combobox with bet slips
Private Sub RS_addComboboxBet(n As String, l As Integer, t As Integer, w As Integer, h As Integer)

    Dim wksR As Worksheet
    Dim strR As Variant
    
    Set collBet = Nothing
    Set collBet = New Collection

    Set comboBet = wks.OLEObjects.Add(classtype:="Forms.ComboBox.1", _
        left:=l, top:=t, width:=w, height:=h)
    
    With comboBet
        .name = n
        .Placement = xlFreeFloating
        .Object.ColumnCount = 2 'column 0: bet slip ID // column 1: visible name
        .Object.ColumnWidths = "20 Pt" 'width of the column with the race name --> hidden
        .Visible = True
    End With
    
'    'Populate Dropdown with installed races
'    For Each wksR In ThisWorkbook.Worksheets
'        If left(wksR.name, 5) = "race_" Then
'            collRennen.Add wksR.name
'            With comboRace.Object
'            .AddItem
'            .List(.ListCount - 1, 0) = wksR.name
'            .List(.ListCount - 1, 1) = wksR.Cells(2, 2).Value & " " & _
'                                    wksR.Cells(3, 2).Value & " (" & _
'                                    wksR.Cells(7, 2).Value & "m)"
'            End With
'        End If
'    Next wksR
 
End Sub

Private Sub RS_HideNavi()
    Dim coll As OLEObject

    For Each coll In wks.OLEObjects
'        Debug.Print "hide... " & coll.name
        coll.Visible = False
    Next coll
    
    wks.Range(Rows(1), Rows(topRows)).Hidden = True
End Sub

Private Sub RS_ShowNavi()
    Dim coll As OLEObject

    wks.Range(Rows(1), Rows(topRows)).Hidden = False

    For Each coll In wks.OLEObjects
'        Debug.Print "show... " & coll.name
        coll.Visible = True
    Next coll

    frmRSnavi.Show

End Sub

Private Sub RS_StartScreen()
    'Prepare Worksheet
        With wks.Range(Columns(1), Columns(80))
            .ColumnWidth = 2
'            .Interior.Color = 16777215
        End With
    
    'Draw Loppsi (get data from worksheet "Pic")
        k = 2
        For i = 1 To 15
            For j = 1 To 41
                wks.Cells(topRows + 1 + i, j).Interior.Color _
                    = wksPic.Cells(k, 2).Value
                k = k + 1
            Next j
        Next i
    'Draw lettering "GALOPPSIM" (get data from worksheet "Pic")
        For i = 1 To 6
            For j = 1 To 41
                wks.Cells(topRows + 1 + i, 26 + j).Interior.Color _
                    = wksPic.Cells(k, 2).Value
                k = k + 1
            Next j
        Next i
    'Draw lettering "RUN SIMPLE"
        k = 2
        For i = 1 To 5
            For j = 1 To 40
                wks.Cells(topRows + 8 + i, 29 + j).Interior.Color _
                    = wksPic.Cells(k, 3).Value
                k = k + 1
            Next j
        Next i
End Sub

Private Sub RS_headings()
    'Menu headings
    wks.Rows(1).Font.name = "Arial Black"
    With wks
        .Cells(1, 2).Value = txt(180)
        .Cells(1, 21).Value = txt(183)
        .Cells(1, 32).Value = txt(181)
        .Cells(1, 60).Value = txt(182)
        .Cells(1, 71).Value = txt(179)
    End With
End Sub

'Grunddaten einlesen
Private Sub Grunddaten1()
    'Fehlerbehandlung
    On Error GoTo FEHLER

    'Prüfen ob es schon ein Tabellenblatt gibt
    For Each wkscheck In ActiveWorkbook.Worksheets
        If wkscheck.name = "GALOPPSIM" Then 'Tabellenblatt ist schon da
            Application.DisplayAlerts = False 'Warnmeldungen ausschalten
            wkscheck.Delete 'Tabellenblatt löschen
            Application.DisplayAlerts = True 'Warnmeldungen einschalten
        End If
    Next wkscheck
    'Neues Tabellenblatt generieren
        Set wks = ActiveWorkbook.Worksheets.Add(before:=Sheets(1)) '(after:=Sheets(Sheets.Count))
        wks.name = "GALOPPSIM"
        wks.Activate

    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Grunddaten einlesen
Private Sub Grunddaten2()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Tabellenblätter zuweisen
    Set wksTxt = ThisWorkbook.Worksheets("Txt")
    Set wksAdv = ThisWorkbook.Worksheets("Adv")
    Set wksPic = ThisWorkbook.Worksheets("Pic")
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Grunddaten einlesen
Private Sub Grunddaten4()
    'Fehlerbehandlung
    On Error GoTo FEHLER

    'Fenster maximieren wenn Checkbox angehakt
        If maxi Then ActiveWindow.WindowState = xlMaximized
    
    'Parameter in Abhängigkeit der Fensterhöhe
        If zoom And ActiveWindow.height >= 600 Then
            spur = 9 'Breite der Rennspuren
            schrift = 8 'Schriftgröße für Pferdenamen und Hufspuren
        Else
            spur = 6 'Breite der Rennspuren
            schrift = 5 'Schriftgröße für Pferdenamen und Hufspuren
        End If
    
    'Parameter in Abhängigkeit der Fensterbreite
        If zoom And ActiveWindow.width >= 1040 Then
            spur2 = 0.3 'Länge der Rennbahn
            liste = 4.5 'Breite Zielfoto/Ergebnistafel
            scrolling = 250 'Scrolling
        Else
            spur2 = 0.3 'Länge der Rennbahn
            liste = 3 'Breite Zielfoto/Ergebnistafel
            scrolling = 200 'Scrolling
        End If
    
    'Geschwindigkeitsspanne des kursfristigen Faktors pro Schleifendurchlauf
        speedK1 = 50300 'Richtwert 50300
        speedK2 = 49700 'Richtwert 49700
        
    'Geschwindigkeitsspanne der Form
        speedF1 = 50010 'Richtwert 50010
        speedF2 = 49990 'Richtwert 49990
    
    'Geschwindigkeiten in den Rennphasen (Taktik)
        speedP1 = 50050 'Richtwert 50050
        speedP2 = 50000
        speedP3 = 49950 'Richtwert 49950
    
    'Spalten bis zu den Boxen (mind. 7, Richtwert 10)
        vorlauf = 11
        
    'Variablen setzen, dass Rennen gestartet wurde
        infoblatt = False
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Grunddaten über das Renneneinlesen
Private Sub Renndaten()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Tabellenblatt mit ausgewähltem Rennen zuweisen
    Set wksData = ThisWorkbook.Worksheets(Auswahl)
    'Grunddaten aus Tabellenblatt einlesen
    With wksData
        raceID = .Cells(1, 2).Value
        race = .Cells(2, 2).Value
        jahr = .Cells(3, 2).Value
        rennort = .Cells(4, 2).Value
        rennbahn = .Cells(5, 2).Value
        Select Case .Cells(6, 2).Value
            Case "T"
                gelaeuf = txt(4)
                trackCol = wksTxt.Cells(4, 6)
            Case "D"
                gelaeuf = txt(5)
                trackCol = wksTxt.Cells(5, 6)
            Case "S"
                gelaeuf = txt(6)
                trackCol = wksTxt.Cells(6, 6)
            Case Else
                gelaeuf = "[error]"
                trackCol = 52377 'green
        End Select
        meter = .Cells(7, 2).Value
        Select Case .Cells(8, 2).Value
            Case "F"
                raceType = txt(170)
                raceTypeID = wksTxt.Cells(170, 6)
            Case "S"
                raceType = txt(171)
                raceTypeID = wksTxt.Cells(171, 6)
            Case Else
                raceType = "[error]"
                raceTypeID = "F" 'flat race
        End Select
        startplaetze = .Cells(9, 2).Value
        randomLane = .Cells(10, 2).Value
        randomCol = .Cells(11, 2).Value
        randomOdd = .Cells(12, 2).Value
        advYN = .Cells(13, 2).Value
    End With
    
    'Advertisement data
    If advYN = "Y" Then
        j = wksData.Cells(Rows.Count, 13).End(xlUp).Row - 1 'last row
        ReDim arrayAdv(1 To j) 'Location of the ad data
        For i = 1 To j
            For k = 1 To wksAdv.Cells(1, Columns.Count).End(xlToLeft).Column
                If wksData.Cells(i + 1, 13).Value = wksAdv.Cells(1, k).Value Then
                    arrayAdv(i) = k 'assign column number
                    Exit For
                End If
            Next k
        Next i
    End If
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Texte einlesen
Private Sub Texte()
    Select Case language
        Case "DE"
            j = 2
        Case "EN"
            j = 3
        Case "RU"
            j = 4
    End Select
    
    For i = 1 To UBound(txt)
        txt(i) = wksTxt.Cells(i, j).Value
    Next i
End Sub

'Daten über die Pferde
Private Sub Pferdedaten()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Anzahl der Starter aus Tabellenblatt auslesen
        starter = Application.WorksheetFunction.CountIf(wksData.Columns(6), "START")
    'Arrays anlegen
    ReDim arrayPferde(1 To startplaetze, 0 To 18) 'Alle Daten der Pferde
    ReDim arrayFotofinish(1 To startplaetze, 0 To 1) 'Snapshot für Zielfoto/Fotofinish
    ReDim arrayErgebnisse(0 To starter, 0 To 5) 'Ergebnisliste
    
    'In case of random lanes
        If randomLane = "R" Then
            Dim boxNr As Integer
            Dim inBox As Boolean
            Dim boxenNr() As Integer
            ReDim boxenNr(1 To startplaetze)
            For i = 1 To startplaetze
                boxenNr(i) = i
            Next i
        End If
    
    For i = 1 To startplaetze
        arrayPferde(i, 0) = wksData.Cells(1 + i, 6).Value 'Status (START, CANCELLED...)
        arrayPferde(i, 11) = wksData.Cells(1 + i, 5).Value 'Startnummer
        arrayPferde(i, 1) = wksData.Cells(1 + i, 7).Value 'Name des Pferds
        If randomCol = "F" Then 'fix
            arrayPferde(i, 2) = wksData.Cells(1 + i, 8).Value 'Horse colour
        Else 'random
            If arrayPferde(i, 1) = "Loppsi" Then
                arrayPferde(i, 2) = 192 'Loppsi is always red
            Else
                Randomize 'Zufallsgenerator zurücksetzen
                Do
                    arrayPferde(i, 2) = CLng(((10 - 1 + 1) * Rnd + 1) * 1000000)
                Loop Until arrayPferde(i, 2) >= 0 And arrayPferde(i, 2) <= 16777215 'allowed value range
            End If
        End If
        
        If randomLane = "R" Then 'lanes random
            inBox = False
            Do Until inBox = True
                Randomize
                boxNr = (Int((startplaetze - 1 + 1) * Rnd + 1)) 'Zufallszahl
                If boxenNr(boxNr) <> 0 Then
                    arrayPferde(i, 15) = boxNr 'Box aus der das Pferd startet
                    boxenNr(boxNr) = 0
                    inBox = True
                End If
            Loop
        Else 'lanes fix
            arrayPferde(i, 15) = wksData.Cells(1 + i, 4).Value 'Box aus der das Pferd startet
        End If
        
        arrayPferde(i, 3) = 5 + 2 * arrayPferde(i, 15) 'Zeilennummer auf der das Pferd läuft
        
        arrayPferde(i, 4) = vorlauf + 5 'Fixe Startposition in der Box
        arrayPferde(i, 5) = wksData.Cells(1 + i, 9).Value 'Grundgeschwindigkeit des Pferds
                                                        '(linear von 1,50010 bis 1,49990)
        'Form des Pferds durch Zufallszahl festlegen
            Randomize 'Zufallsgenerator zurücksetzen
            arrayPferde(i, 6) = (Int((speedF1 - speedF2 + 1) * Rnd + speedF2) + 100000) / 100000 'Zufallszahl
        'Wettquote festlegen
            If randomOdd = "F" Then 'fix
                arrayPferde(i, 17) = wksData.Cells(1 + i, 10).Value
            Else 'random
                Randomize 'Zufallsgenerator zurücksetzen
                Do
                    arrayPferde(i, 17) = Round(CInt((1 + (((Int((4 - 0 + 1) * Rnd + 0)) - 2) / 10)) _
                        * (150012 - (arrayPferde(i, 6) * 100000)) * 10) / 5) * 5 'complicated formula
                Loop Until arrayPferde(i, 17) >= 15 'minimum value
            End If
        'Schätzfehler für Balkenanzeige bei Wetten (+/-50)
            Randomize 'Zufallsgenerator zurücksetzen
            arrayPferde(i, 18) = (Int((100 - 0 + 1) * Rnd + 0)) - 50 'random number between -50 and +50
    Next i
    
    'Favoriten errmitteln aus Grundgeschwindigkeit und Form
        'Variablen zurücksetzen
            favsum1 = 0
            favsum2 = 0
            favsum3 = 0
        'Berechnung der drei Favoriten
            For i = 1 To startplaetze
                If arrayPferde(i, 0) = "START" Then
                    If arrayPferde(i, 5) + arrayPferde(i, 6) > favsum1 Then
                        favsum3 = favsum2
                        favsum2 = favsum1
                        favsum1 = arrayPferde(i, 5) + arrayPferde(i, 6)
                        favorit3 = favorit2
                        favorit2 = favorit1
                        favorit1 = i 'Nummer des Favoriten
                    ElseIf arrayPferde(i, 5) + arrayPferde(i, 6) > favsum2 Then
                        favsum3 = favsum2
                        favsum2 = arrayPferde(i, 5) + arrayPferde(i, 6)
                        favorit3 = favorit2
                        favorit2 = i 'Nummer eines weiteren Favoriten
                    ElseIf arrayPferde(i, 5) + arrayPferde(i, 6) > favsum3 Then
                        favsum3 = arrayPferde(i, 5) + arrayPferde(i, 6)
                        favorit3 = i 'Nummer eines weiteren Favoriten
                    End If
                End If
            Next i
        
        'Favoriten in Array eintragen
            arrayPferde(favorit1, 16) = 1
            arrayPferde(favorit2, 16) = 2
            arrayPferde(favorit3, 16) = 3
    
    'Wenn Checkbox angehakt: Geschwindigkeit im 1., 2. und 3. Renndrittel pro Pferd festlegen
        If taktik Then
            For i = 1 To startplaetze
                Randomize 'Zufallsgenerator zurücksetzen
                Select Case (Int((6 - 1 + 1) * Rnd + 1)) 'Zufallszahl zwischen 1 und 6
                    Case 1
                        arrayPferde(i, 12) = (speedP1 + 100000) / 100000
                        arrayPferde(i, 13) = (speedP2 + 100000) / 100000
                        arrayPferde(i, 14) = (speedP3 + 100000) / 100000
                    Case 2
                        arrayPferde(i, 12) = (speedP1 + 100000) / 100000
                        arrayPferde(i, 13) = (speedP3 + 100000) / 100000
                        arrayPferde(i, 14) = (speedP2 + 100000) / 100000
                    Case 3
                        arrayPferde(i, 12) = (speedP2 + 100000) / 100000
                        arrayPferde(i, 13) = (speedP1 + 100000) / 100000
                        arrayPferde(i, 14) = (speedP3 + 100000) / 100000
                    Case 4
                        arrayPferde(i, 12) = (speedP3 + 100000) / 100000
                        arrayPferde(i, 13) = (speedP1 + 100000) / 100000
                        arrayPferde(i, 14) = (speedP2 + 100000) / 100000
                    Case 5
                        arrayPferde(i, 12) = (speedP2 + 100000) / 100000
                        arrayPferde(i, 13) = (speedP3 + 100000) / 100000
                        arrayPferde(i, 14) = (speedP1 + 100000) / 100000
                    Case 6
                        arrayPferde(i, 12) = (speedP3 + 100000) / 100000
                        arrayPferde(i, 13) = (speedP2 + 100000) / 100000
                        arrayPferde(i, 14) = (speedP1 + 100000) / 100000
                End Select
            Next i
        End If
        
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

Private Sub Rennstrecke()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Bildschirmaktualisierung ausschalten
        Application.ScreenUpdating = False
        
    'Spalten A-D des Fensters fixieren wenn Checkbox für Namen oder Farbe der Pferde
    'oder die Markierung des Favoriten am linken Rand angehakt ist
        If namen Or pferdefarbe Or markFav Then
            With ActiveWindow
                .SplitColumn = 5
                .SplitRow = 0
                .FreezePanes = True
            End With
        End If
    'Gitternetz anzeigen wenn Checkbox angehakt
        If gitter Then
            ActiveWindow.DisplayGridlines = False
        Else
            ActiveWindow.DisplayGridlines = True
        End If
    'Zeilen-/Spaltennummern anzeigen wenn Checkbox angehakt
        If headings Then
            ActiveWindow.DisplayHeadings = False
        Else
            ActiveWindow.DisplayHeadings = True
        End If
    'Zeilenhöhen
        wks.Range(Rows(1 + topRows), Rows(5 + topRows)).EntireRow.RowHeight = 15 'oberhalb der Rennbahn
        wks.Range(Rows(startplaetze * 2 + 6 + 1 + topRows), _
            Rows(startplaetze * 2 + 52 + topRows)).EntireRow.RowHeight = 15 'unterhalb der Rennbahn
        wks.Rows(startplaetze * 2 + 20 + topRows).RowHeight = 20 'Überschrift Ergebnistafel
    'Renndaten anzeigen
        With wks.Cells(2 + topRows, 6)
            .Font.name = "Arial Black"
            .Value = race & " " & jahr & " - " & rennbahn & ", " & rennort
        End With
        With wks.Cells(3 + topRows, 6)
            .Font.name = "Arial"
            .Font.Bold = True 'Fettschrift
            .Value = raceType & " " & txt(1) & " " & _
                meter & txt(2) & " - " & gelaeuf
        End With
    'Bereich vor Startboxen
        wks.Columns(1).ColumnWidth = 1 'Linker Rand
        wks.Columns(2).ColumnWidth = 1.5 'Spalte für Farbe der Pferde
        wks.Columns(3).ColumnWidth = 1 'Leere Spalte
        wks.Columns(4).ColumnWidth = 4.5 'Spalte für Startnummern
        wks.Columns(5).ColumnWidth = 22 'Spalte für Pferdenamen
        wks.Range(Columns(6), Columns(vorlauf + 5)).ColumnWidth = spur2
        wks.Columns(vorlauf - 3).ColumnWidth = 6 'Spalte für Boxennummern
    'Rennbahn
        wks.Range(Columns(vorlauf + 6), Columns(meter + vorlauf + 6)).ColumnWidth = spur2 'Länge der Bahn
        wks.Range(Rows(6 + topRows), Rows(startplaetze * 2 + 6 + topRows)).RowHeight = spur 'Spurbreite
        wks.Range(Cells(4 + topRows, 1), Cells(startplaetze * 2 + 19 + topRows, meter + vorlauf + 7)).Interior.Color = trackCol
    'Schriftart
        wks.Range(Cells(4 + topRows, 1), Cells(startplaetze * 2 + 8 + topRows, meter + vorlauf + 7)).Font.name = "Arial"
    'Startboxen
        For i = 6 To (startplaetze * 2 + 6) Step 2 'Für jeden Startplatz eine Box
            wks.Range(Cells(i + topRows, vorlauf + 1), Cells(i + topRows, vorlauf + 6)).Interior.ColorIndex = 1 'schwarz
        Next i
        wks.Range(Cells(6 + topRows, vorlauf + 6), Cells(startplaetze * 2 + 6 + topRows, vorlauf + 6)).Interior.ColorIndex = 1 'schwarz
    'Beschriftung der Boxen
        With wks.Range(Cells(7 + topRows, vorlauf - 3), Cells(startplaetze * 2 + 5 + topRows, vorlauf - 3))
            .Font.ColorIndex = 1 'schwarz
            .Font.Size = schrift
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        For i = 1 To (startplaetze) 'Für jeden Startplatz eine Box
            wks.Cells(5 + 2 * i + topRows, vorlauf - 3).Value = txt(8) & i
        Next i
    'Meter-Anzeigen
        For i = 250 To meter - 250 Step 250
'            wks.Range(Cells(5, i + vorlauf + 5), Cells(45, i + vorlauf + 5)).Interior.ColorIndex = 1 'für Entwicklung einkommentieren
            With wks.Cells(4 + topRows, i + vorlauf)
                .Font.name = "Arial"
                .Font.Bold = True 'Fettschrift
                .Value = i & txt(2) '"m"
            End With
            With wks.Cells(startplaetze * 2 + 8 + topRows, i + vorlauf)
                .Font.name = "Arial"
                .Font.Bold = True 'Fettschrift
                .Value = i & txt(2) '"m"
            End With
        Next i
    'Formatierung für Pferdenamen am linken Rand
        With wks.Range(Cells(6 + topRows, 4), Cells(startplaetze * 2 + 7 + topRows, 5))
            .Font.Color = trackCol  'wie track, damit zuerst nicht sichtbar
            .IndentLevel = 1 'Text eingerückt
            .Font.Size = schrift
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    '"Schriftgröße" für Hufspuren
        With wks.Range(Cells(5 + topRows, vorlauf - 1), Cells(startplaetze * 2 + 7 + topRows, meter + vorlauf))
            .Font.Size = schrift
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
    'Ziel
        wks.Columns(meter + vorlauf + 5).ColumnWidth = spur2 'Ziellinie
        wks.Range(Cells(5 + topRows, meter + vorlauf + 5), _
            Cells(startplaetze * 2 + 7 + topRows, meter + vorlauf + 5)).Interior.ColorIndex = 56 'grau
        With wks.Cells(4 + topRows, meter + vorlauf + 1)
            .Font.name = "Arial"
            .Font.Bold = True 'Fettschrift
            .Value = meter & txt(2) 'Meter im Ziel
        End With
        With wks.Cells(startplaetze * 2 + 8 + topRows, meter + vorlauf + 1)
            .Font.name = "Arial"
            .Font.Bold = True 'Fettschrift
            .Value = meter & txt(2)  'Meter im Ziel
        End With
    'Auslauf hinter Ziel
        wks.Columns(meter + vorlauf + 7).ColumnWidth = 30
        wks.Columns(meter + vorlauf + 8).ColumnWidth = 10
    'Formatierung für Pferdenamen im Ziel
        With wks.Range(Cells(5 + topRows, meter + vorlauf + 7), Cells(startplaetze * 2 + 7 + topRows, meter + vorlauf + 7))
            .Font.ColorIndex = 1 'schwarz
            .IndentLevel = 1 'Text eingerückt
            .Font.Size = schrift
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    
    'Werbung
        If advYN = "Y" Then
            wks.Range(Rows(startplaetze * 2 + 9 + topRows), _
                Rows(startplaetze * 2 + 19 + topRows)).EntireRow.RowHeight = 5 'unterhalb der Rennbahn
            Dim advPos As Integer 'column position for next ad
            advPos = vorlauf + 5
            For i = 1 To UBound(arrayAdv)
                z = 3 'set pointer to first colour code
                For j = advPos To advPos + wksAdv.Cells(2, arrayAdv(i)) - 1
                    If j >= meter + vorlauf + 5 Then Exit For
                    For k = startplaetze * 2 + 9 + topRows To startplaetze * 2 + 19 + topRows
                        wks.Cells(k, j).Interior.Color = wksAdv.Cells(z, arrayAdv(i)).Value
                        z = z + 1
                    Next k
                Next j
                advPos = advPos + wksAdv.Cells(2, arrayAdv(i))
            Next i
        End If
    
    'Formatierungen für Zielfoto
        wks.Cells(2 + topRows, meter + vorlauf + 7).Font.name = "Arial Black" '"FOTOFINISH!"
        With wks.Cells(4 + topRows, meter + vorlauf + 9) '"Zielfoto..."
'            .Font.name = "Courier New"
            .Font.Size = 14
            .Font.Bold = True 'Fettschrift
        End With
    'Formatierung für Ergebnisliste/Zielfoto
        wks.Range(Columns(meter + vorlauf + 9), Columns(meter + vorlauf + 31)).ColumnWidth = liste
    'Formatierung für Siegerfoto
        wks.Range(Columns(meter + vorlauf + 28), Columns(meter + vorlauf + 44)).ColumnWidth = 2
        With wks.Range(Cells(startplaetze * 2 + 20 + topRows, meter + vorlauf + 29), _
                        Cells(startplaetze * 2 + 21 + topRows, meter + vorlauf + 31))
'            .Font.name = "Courier New"
            .Font.Size = 14
            .Font.Bold = True 'Fettschrift
        End With
    'Formatierung für Tool-Infos
        Call ToolInfoFormatierung
    'Cursor weit weg platzieren
        wks.Cells(100 + topRows, 1).Select
    'Bildschirmaktualisierung ausschalten
        Application.ScreenUpdating = True
        
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Pferde generieren
Private Sub Startfeld()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Pferdenamen am linken Rand platzieren wenn Checkbox angehakt ist
        If namen Then
            For i = 1 To startplaetze
                If arrayPferde(i, 0) = "START" Then
                    wks.Cells(arrayPferde(i, 3) + topRows, 4).Value = "#" & arrayPferde(i, 11) 'Startnummer
                    wks.Cells(arrayPferde(i, 3) + topRows, 5).Value = arrayPferde(i, 1) 'Name des Pferds
                End If
            Next i
        'Optimale Spaltenbreite
            wks.Range(Columns(3), Columns(4)).EntireColumn.AutoFit
        End If
    'Pferdenamen im Ziel anzeigen wenn Checkbox angehakt ist
        If namen2 Then
            For j = 1 To startplaetze
                If arrayPferde(j, 0) = "START" Then
                    wks.Cells(arrayPferde(j, 3) + topRows, meter + vorlauf + 7).Value = _
                        arrayPferde(j, 1) & " (#" & arrayPferde(j, 11) & ")"
                End If
            Next j
        End If
        
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Meldung zum Rennstart
Private Sub Begruessung()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    MsgBox txt(9) & " " & txt(14) & " " & rennort & ". " & vbNewLine & _
            txt(11) & " " & race & " " & txt(1) & " " & meter & " " & txt(3) & "." & vbNewLine & _
            txt(12) & " " & starter & " " & txt(13), , tool

    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Pferde in Boxen stellen
Private Sub Startaufstellung()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    If test = False Then Application.Wait (Now + TimeValue("0:00:02")) 'Verzögerung
    For i = 1 To startplaetze
        If arrayPferde(i, 0) = "START" Then
            wks.Range(Cells(arrayPferde(i, 3) + topRows, arrayPferde(i, 4)), _
                Cells(arrayPferde(i, 3) + topRows, arrayPferde(i, 4) - 7)) _
                .Interior.Color = arrayPferde(i, 2)
        End If
    Next i
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Vorstellung der Pferde mit Startnummer und Namen
Private Sub Vorstellung()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    If test = False Then Application.Wait (Now + TimeValue("0:00:02")) 'Verzögerung
    Application.DisplayCommentIndicator = xlCommentAndIndicator 'Kommentare anschalten
    
    For i = 1 To startplaetze
        If arrayPferde(i, 0) = "START" Then
            With wks.Cells(arrayPferde(i, 3) + topRows, arrayPferde(i, 4))
                .AddComment Text:="#" & arrayPferde(i, 11) & " " & arrayPferde(i, 1)
                .Comment.Shape.TextFrame.AutoSize = True
            End With
            If i = favorit1 Then
                wks.Cells(arrayPferde(i, 3) + topRows, arrayPferde(i, 4)) _
                    .Comment.Shape.Fill.ForeColor.RGB = RGB(192, 0, 0)
            End If
        End If
    Next i
    
    If test = False Then Application.Wait (Now + TimeValue("0:00:02")) 'Verzögerung
    'Messagebox
        MsgBox txt(15) & " " & arrayPferde(favorit1, 1) & _
                " " & txt(18) & " " & arrayPferde(favorit1, 15) & _
                "." & vbNewLine & vbNewLine & _
                txt(17) & " " & arrayPferde(favorit2, 1) & " " & txt(18) & " " & _
                arrayPferde(favorit2, 15) & vbNewLine & _
                txt(19) & " " & _
                arrayPferde(favorit3, 1) & " " & txt(18) & " " & _
                arrayPferde(favorit3, 15) & " " & txt(20) & ".", , tool
    'Kommentare ausblenden
        Application.DisplayCommentIndicator = xlNoIndicator
    'Farben der Pferde am linken Rand platzieren wenn Checkbox angehakt ist
        If pferdefarbe Then
            For i = 1 To startplaetze
                If arrayPferde(i, 0) = "START" Then
                    wks.Cells(arrayPferde(i, 3) + topRows, 2).Interior.Color = arrayPferde(i, 2) 'Farbe des Pferds
                End If
            Next i
        End If
    'Pferdenamen und Startnummern am linken Rand zeigen wenn Checkbox angehakt ist
        wks.Range(Columns(4), Columns(5)).Font.ColorIndex = 1 'schwarz
    
    'Favoriten markieren wenn angehakt
        If markFav Then
            wks.Range(Cells(arrayPferde(favorit1, 3) + topRows, 4), Cells(arrayPferde(favorit1, 3) + topRows, 5)) _
                .Interior.Color = 192
        End If
    
    'Verzögerung
        If test = False Then Application.Wait (Now + TimeValue("0:00:04"))

    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Start des Galopprennens
Private Sub Rennstart()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Boxennummern entfernen
        wks.Range(Cells(7 + topRows, vorlauf - 3), Cells(5 + 2 * startplaetze + topRows, vorlauf - 3)).Value = ""
    'Verzögerung
        If test = False Then Application.Wait (Now + TimeValue("0:00:04"))
    'Boxen aufmachen
        wks.Range(Cells(6 + topRows, vorlauf + 6), Cells(startplaetze * 2 + 6 + topRows, vorlauf + 6)).Interior.Color = trackCol
    'Fotofinish zurücksetzen
        fotofin = 0
    'Noch kein Pferd ist im Ziel
        sieg = False
    'Platzierung für das nächste Pferd, das im Ziel ankommt
        platz = 1
    'Rennen läuft
        Do Until platz > starter 'solange noch nicht alle im Ziel sind
            'Zählvariable für den Zieleinlauf pro Schleifendurchlauf zurücksetzen
                einlauf = 0
            'Neue Positionen berechnen
                For i = 1 To UBound(arrayPferde)
                    'Geschwindigkeitsfaktor pro Durchlauf
                    arrayPferde(i, 7) = speedKurz()
                    'Schrittweite pro Durchlauf (ungerundet)
                        If taktik = False Then
                        'Wenn Geschwindigkeit pro Rennphase konstant sein soll
                            arrayPferde(i, 8) = _
                                (arrayPferde(i, 5) + arrayPferde(i, 6) + arrayPferde(i, 7)) / 3
                        Else
                        'Wenn jedes Pferd in einem Renndrittel unterschiedlich schnell sein soll
                            'Berechnen in welchem Streckenabschnitt das Pferd ist
                            Select Case True
                                Case arrayPferde(i, 4) < meter * 1 / 3 'Pferd ist im 1. Renndrittel
                                    arrayPferde(i, 8) = _
                                        (arrayPferde(i, 5) + arrayPferde(i, 6) + _
                                            arrayPferde(i, 7) + arrayPferde(i, 12)) / 4
                                Case arrayPferde(i, 4) < meter * 2 / 3 'Pferd ist im 3. Renndrittel
                                    arrayPferde(i, 8) = _
                                        (arrayPferde(i, 5) + arrayPferde(i, 6) + _
                                            arrayPferde(i, 7) + arrayPferde(i, 13)) / 4
                                Case Else 'Pferd ist im 3. Renndrittel
                                    arrayPferde(i, 8) = _
                                        (arrayPferde(i, 5) + arrayPferde(i, 6) + _
                                            arrayPferde(i, 7) + arrayPferde(i, 14)) / 4
                                End Select
                        End If
                    'Schrittweite pro Durchlauf (1 oder 2 Spalten)
                    arrayPferde(i, 9) = Round(arrayPferde(i, 8), 0)
                Next i
            'Pferde laufen
                For i = 1 To UBound(arrayPferde)
                    If arrayPferde(i, 0) <> "CANCELLED" Then 'nur wenn Pferd am Start ist
                        'Pferd löschen
                            Range(Cells(arrayPferde(i, 3) + topRows, arrayPferde(i, 4)), _
                                Cells(arrayPferde(i, 3) + topRows, arrayPferde(i, 4) - 7)) _
                                .Interior.Color = trackCol
                        'Neue Position des Pferds festlegen (nur wenn Pferd noch läuft)
                            If arrayPferde(i, 0) = "START" Then
                                arrayPferde(i, 4) = arrayPferde(i, 4) + arrayPferde(i, 9)
                            End If
                        'Pferd neu setzen (auch die, die schon im Ziel sind wegen dem Rendering)
                            Range(Cells(arrayPferde(i, 3) + topRows, arrayPferde(i, 4)), _
                                Cells(arrayPferde(i, 3) + topRows, arrayPferde(i, 4) - 7)) _
                                .Interior.Color = arrayPferde(i, 2)
                        'Wenn Checkbox angehakt ist: Hufspur zeichnen
                            If hufspur Then Cells(arrayPferde(i, 3) + topRows, arrayPferde(i, 4) - 8).Value = "."
                    End If
                    'Horizontal scrollen wenn nötig
                    If arrayPferde(i, 4) > ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.Count).Column - 50 _
                        And ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.Count).Column <= meter + 11 Then
                            'Scrollen
                            ActiveWindow.ScrollColumn = ActiveWindow.VisibleRange.Column + scrolling
                            If ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.Count).Column + 30 >= meter + vorlauf Then
                                ActiveWindow.ScrollColumn = ActiveWindow.VisibleRange.Column + 30
                            End If 'ToDo... is this correct?
                    End If
                Next i
            'Check ob ein Pferd im Ziel ist
                For i = 1 To UBound(arrayPferde)
                    If arrayPferde(i, 0) = "START" Then 'nur wenn Pferd noch läuft
                        If arrayPferde(i, 4) >= meter + vorlauf + 5 Then 'Ziellinie erreicht
                            arrayPferde(i, 0) = "BERECHNUNG"
                            einlauf = einlauf + 1 'zählen, wie viele Pferde in diesem Durchlauf ins Ziel kommen
                        End If
                    End If
                Next i
                
                If einlauf > 0 Then
                    If sieg = False Then
                        If einlauf > 1 Then
                            fotofin = True 'Fotofinish (Zielfoto s/w)
                            'Text anpassen
                                wks.Cells(2 + topRows, meter + vorlauf + 7).Value = txt(21) '"FOTOFINISH!"
                        End If
                        Call Zielfoto 'Zielfoto machen
                    End If
                    sieg = True 'damit das Zielfoto nur 1x gemacht wird
                    Call Platzierung 'Absprung in die Platzberechnung, wenn mehr als ein Pferd in dieser Runde ins Ziel kommen
                End If
            DoEvents 'Rendering
        Loop
    
    'Wenn es ein Fotofinish gibt
        If fotofin = True Then
            'Verzögerung
                If test = False Then Application.Wait (Now + TimeValue("0:00:04"))
            'Texte anpassen
                wks.Cells(2 + topRows, meter + vorlauf + 7).Value = ""
                wks.Cells(4 + topRows, meter + vorlauf + 9).Value = txt(22) '("Zielfoto wird erstellt")
            'Scrollen
                On Error Resume Next
                ActiveWindow.ScrollColumn = meter - 20
                On Error GoTo 0
            'Verzögerung
                If test = False Then Application.Wait (Now + TimeValue("0:00:04"))
            'Zielfoto anzeigen
                Call FotoZeigen
            'Text anpassen
                wks.Cells(4 + topRows, meter + vorlauf + 9).Value = ""
                wks.Cells(4 + topRows, meter + vorlauf + 9).Value = txt(23) '("Zielfoto wird ausgewertet")
            'Verzögerung
                If test = False Then Application.Wait (Now + TimeValue("0:00:04"))
            'Text anpassen
                wks.Cells(4 + topRows, meter + vorlauf + 9).Value = txt(28) '("Zielfoto")
        End If

    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Platzierung berechnen wenn ein oder mehrere Pferde in einem Schleifendurchlauf ins Ziel kommen
Private Sub Platzierung()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    ReDim arrayBerechnung(1 To einlauf, 0 To 4)
    
    Dim zuf1 As Integer, zuf2 As Double 'Variablen für Zufallszahlen
    Dim aktuelleRunde As Boolean 'Kennzeichen, dass eine neue Berechnungsrunde läuft
    Dim p As Boolean 'Platzierung wird vergeben wenn TRUE
    Dim fertig As Integer 'Berechnung fertig wenn alle Plätze vergeben sind
    Dim m As Integer 'Zählvariable

    aktuelleRunde = False
    p = False
    fertig = 0
    m = 1

    'Positionen in Berechnungs-Array eintragen
    For i = 1 To UBound(arrayPferde)
        If arrayPferde(i, 0) = "BERECHNUNG" Then
            arrayPferde(i, 0) = "ZIEL" 'Finalen Status setzen
            arrayBerechnung(m, 1) = arrayPferde(i, 11) 'Startnummer
            arrayBerechnung(m, 2) = arrayPferde(i, 1) 'Name des Pferds
            arrayBerechnung(m, 3) = arrayPferde(i, 4) 'Position des Pferds
            arrayBerechnung(m, 4) = arrayPferde(i, 2) 'Farbe des Pferds
            m = m + 1
        End If
    Next i

    'Exakte Position des Pferds per Zufallszahlen generieren
    For i = 1 To UBound(arrayBerechnung)
        'Position durch Zufall neu berechnen
            Randomize 'Zufallsgenerator zurücksetzen
            zuf1 = (Int((2 - 1 + 1) * Rnd + 1)) '1 = addieren, 2 = subtrahieren
            Randomize 'Zufallsgenerator zurücksetzen
            If zuf1 = 1 Then
                arrayBerechnung(i, 3) = Round(arrayBerechnung(i, 3) _
                    + (Int((5 - 1 + 1) * Rnd + 1) / 10), 1) 'Dezimalstellen x.1 bis x.5
            Else
                arrayBerechnung(i, 3) = Round(arrayBerechnung(i, 3) _
                    - (Int((4 - 0 + 0) * Rnd + 0.5) / 10), 1) 'Dezimalstellen (x-1).6 bis x
            End If
    Next i
    
    'Platzierungen vergeben
    Do Until fertig >= UBound(arrayBerechnung)
        For i = 1 To UBound(arrayBerechnung)
            If arrayBerechnung(i, 0) <> "PLATZIERT" Then
                For j = i To UBound(arrayBerechnung)
                    If arrayBerechnung(j, 0) <> "PLATZIERT" Then
                        If arrayBerechnung(i, 3) >= arrayBerechnung(j, 3) Then 'Position ist größer als die des Vergleichspferds
                                p = True
                        Else 'Vergleichspferd ist weiter vorne
                            p = False
                            Exit For
                        End If
                    End If
                Next j
                If p = True Then
                    arrayBerechnung(i, 0) = "PLATZIERT" 'eintragen, dass das Pferd nicht mehr verglichen wird
                    aktuelleRunde = True
                    fertig = fertig + 1 'hochzählen
                    p = False 'zurücksetzen
                    'Pferd in Ergebnisliste eintragen
                        arrayErgebnisse(platz, 2) = arrayBerechnung(i, 1) 'Startnummer
                        arrayErgebnisse(platz, 3) = arrayBerechnung(i, 2) 'Name des Pferds
                        arrayErgebnisse(platz, 4) = arrayBerechnung(i, 4) 'Farbe des Pferds
                        arrayErgebnisse(platz, 5) = arrayBerechnung(i, 3) 'Position des Pferds
                        'Platzierung berechnen
                            If aktuelleRunde And _
                                arrayErgebnisse(platz, 5) = arrayErgebnisse(platz - 1, 5) Then
                                    'wenn exakt gleich wie das Pferd zuvor in dieser Berechnungsrunde
                                    arrayErgebnisse(platz, 1) = platz - 1
                            Else
                                'wenn Position kleiner ist als beim Pferd zuvor
                                arrayErgebnisse(platz, 1) = platz
                            End If
                        platz = platz + 1 'Platz für nächstes Pferd hochzählen
                    Exit For
                End If
            End If
        Next i
    Loop
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Zielfoto machen wenn das erste Pferd das Ziel erreicht
Private Sub Zielfoto()
    'Fehlerbehandlung
    On Error GoTo FEHLER
            
    'Daten eintragen
        For j = 1 To UBound(arrayFotofinish)
            arrayFotofinish(j, 0) = arrayPferde(j, 3) 'Spur
            arrayFotofinish(j, 1) = arrayPferde(j, 4) 'Position des Pferds
        Next j
    'Blitz wenn Fotofinish
        If fotofin Then
            For k = 1 To 8
                With wks.Range(Cells(5 + topRows, meter + vorlauf + 7), _
                    Cells(startplaetze * 2 + 7 + topRows, meter + vorlauf + 7))
                        .Interior.ColorIndex = 1 'schwarz
                        .Interior.ColorIndex = 0 'weiß
                End With
            Next k
            wks.Range(Cells(5 + topRows, meter + vorlauf + 7), _
                Cells(startplaetze * 2 + 7 + topRows, meter + vorlauf + 7)).Interior.Color = trackCol
        End If

    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Zielfoto anzeigen
Private Sub FotoZeigen()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Zielfoto zeichnen
        Application.ScreenUpdating = False 'Bildschirmaktualisierung ausschalten
        If fotofin = True Then
            'Geläuf und Ziellinie (s/w)
            wks.Range(Cells(5 + topRows, meter + vorlauf + 9), Cells(startplaetze * 2 + 7 + topRows, meter + vorlauf + 24)).Interior.ColorIndex = 1 'Rasen schwarz
            wks.Range(Cells(5 + topRows, meter + vorlauf + 22), Cells(startplaetze * 2 + 7 + topRows, meter + vorlauf + 22)).Interior.ColorIndex = 0 'Zielline weiß
        Else
            'Geläuf und Ziellinie (Originalfarben)
            wks.Range(Cells(5 + topRows, meter + vorlauf + 9), Cells(startplaetze * 2 + 7 + topRows, meter + vorlauf + 24)).Interior.Color = trackCol
            wks.Range(Cells(5 + topRows, meter + vorlauf + 22), Cells(startplaetze * 2 + 7 + topRows, meter + vorlauf + 22)).Interior.ColorIndex = 56 'Zielline grau
        End If
    'Pferde zeichnen
        Dim p As Integer 'Pferd auf Zielfoto
        For i = 1 To UBound(arrayFotofinish)
            If arrayFotofinish(i, 1) >= meter + vorlauf + 5 - 13 Then 'nur wenn Pferd im Foto ist
                If arrayFotofinish(i, 1) - 7 >= meter + vorlauf + 5 - 13 Then
                    p = arrayFotofinish(i, 1) - 7
                Else
                    p = meter + 3
                End If
                Range(Cells(arrayFotofinish(i, 0) + topRows, arrayFotofinish(i, 1) + 17), _
                    Cells(arrayFotofinish(i, 0) + topRows, p + 17)) _
                    .Interior.Color = arrayPferde(i, 2) 'Pferd setzen
            End If
        Next i
    'Bildschirmaktualisierung einschalten
        Application.ScreenUpdating = True
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Ergebnisliste anzeigen
Private Sub Ergebnisse()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Verzögerung
        If test = False Then Application.Wait (Now + TimeValue("0:00:02"))

    MsgBox txt(24) & vbNewLine & txt(25), , tool
    'Fixierung des Fensters wieder aufheben wenn Checkbox für Namen oder Farbe der Pferde
    'oder die Markierung des Favoriten am linken Rand angehakt ist
        If namen Or pferdefarbe Or markFav Then
            With ActiveWindow
                .SplitColumn = 0
                .SplitRow = 0
                .FreezePanes = False
            End With
        End If
    'Scrollen zu Ergebnistafel
        Call ScrollErgebnis(startplaetze * 2 + 19)
    'Anzeigetafel
        With wks.Range(Cells(startplaetze * 2 + 20 + topRows, meter + vorlauf + 9), _
            Cells(startplaetze * 2 + 20 + starter + 1 + topRows, meter + vorlauf + 27))
                .Interior.Color = 16777215 'Hintergrund
                .Font.name = "Courier New"
                .Font.Size = 12
                .NumberFormat = "@" 'Textformat
        End With
        With wks.Cells(startplaetze * 2 + 20 + topRows, meter + vorlauf + 9) 'Überschrift
            .Font.Size = 14
            .Font.Bold = True 'Fettschrift
            .IndentLevel = 1 'Text eingerückt
        End With

    'Rahmen um die Anzeigetafel
        wks.Range(Cells(startplaetze * 2 + 20 + topRows, meter + vorlauf + 9), _
            Cells(startplaetze * 2 + 20 + starter + 1 + topRows, meter + vorlauf + 27)).BorderAround ColorIndex:=0
    
    'Ergebnisse eintragen
        'Überschrift
            wks.Cells(startplaetze * 2 + 20 + topRows, meter + vorlauf + 9).Value _
                = race & " " & jahr
        'Verzögerung
            If spannung Then Application.Wait (Now + TimeValue("0:00:04"))
        'Platzierungen anzeigen
            Dim posH As Integer 'Position of the horse names
            posH = 0
            For i = UBound(arrayErgebnisse) To 1 Step -1
                If colResult Then
                    posH = 3
                    Range(Cells(startplaetze * 2 + 20 + i + topRows, meter + vorlauf + 10), _
                        Cells(startplaetze * 2 + 20 + i + topRows, meter + vorlauf + 11)) _
                        .Interior.Color = arrayErgebnisse(i, 4) 'Farbe des Pferds
                End If
                Cells(startplaetze * 2 + 20 + i + topRows, meter + vorlauf + 10 + posH).Value = arrayErgebnisse(i, 1) & "." 'Platzierung
                Cells(startplaetze * 2 + 20 + i + topRows, meter + vorlauf + 12 + posH).Value = arrayErgebnisse(i, 3) & _
                    " (#" & arrayErgebnisse(i, 2) & ")" 'Name und Startnummer des Pferds
                'Wenn Checkbox zur Spannungssteigerung angehakt
                If spannung Then Application.Wait (Now + TimeValue("0:00:01")) 'Verzögerung
            Next i

    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Kurzfristige Geschwindigkeit der Pferde (Faktor wird bei jedem Schleifendurchlauf neu berechnet)
Function speedKurz() As Double
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    Randomize 'Zufallsgenerator zurücksetzen
    speedKurz = (Int((speedK1 - speedK2 + 1) * Rnd + speedK2) + 100000) / 100000 'Zufallszahl
    
    Exit Function
FEHLER:
    Call Programmabsturz
End Function

'Scrollen zu Zielfoto, Ergebnistafel oder Wettscheinauswertung
Private Sub ScrollErgebnis(z As Integer)
    On Error Resume Next
        ActiveWindow.ScrollColumn = meter - 20
        ActiveWindow.ScrollRow = z + topRows
    On Error GoTo 0
End Sub

'Scrollen zu Tool-Info
Private Sub ScrollInfo()
    On Error Resume Next
        ActiveWindow.ScrollColumn = meter + vorlauf + 44 + 30 - 10
        ActiveWindow.ScrollRow = startplaetze * 2 + starter + 146 + topRows
    On Error GoTo 0
End Sub

'Pferd mit Siegerkranz zeichnen
Private Sub Siegerpferd()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Verzögerung
        If test = False Then Application.Wait (Now + TimeValue("0:00:01"))
    'Bildschirmaktualisierung ausschalten
        Application.ScreenUpdating = False
    'Pferd zeichnen (auslesen aus Tabellenblatt Pic)
        k = 2
        For i = 1 To 13
            For j = 1 To 18
                wks.Cells(startplaetze * 2 + 22 + i + topRows, meter + vorlauf + 28 + j).Interior.Color _
                    = wksPic.Cells(k, 1).Value
                k = k + 1
            Next j
        Next i

    'Name des Siegers
        sieger = ""
        For i = 1 To UBound(arrayErgebnisse)
            If arrayErgebnisse(i, 1) = 1 Then
                If i > 1 Then sieger = sieger & " und " 'wenn mehrere Pferde gewinnen
                sieger = sieger & UCase(arrayErgebnisse(i, 3))
            End If
        Next i
        wks.Cells(startplaetze * 2 + 20 + topRows, meter + vorlauf + 29).Value = txt(27)
        wks.Cells(startplaetze * 2 + 21 + topRows, meter + vorlauf + 31).Value = sieger
'        wks.Cells(startplaetze * 2 + 20 + topRows, meter + vorlauf + 29).Value = txt(27) & " " & UCase(arrayErgebnisse(1, 3))
    'Bildschirmaktualisierung einschalten
        Application.ScreenUpdating = True
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Analyze bet slips
Private Sub AnalyseBetSlips()
    Dim id As String
    Dim nm As String
    Dim st As Double
    Dim od As Double
    Dim bt() As Integer
    Dim payout As Boolean
    Dim noWinner As Boolean
    
    If collSlips.Count > 0 Then
    
    noWinner = True 'Reset variable
        
        'Anzeigetafel
        With wks.Range(Cells(startplaetze * 2 + 20 + starter + 4 + topRows, meter + vorlauf + 9), _
            Cells(startplaetze * 2 + 20 + starter + collSlips.Count + 4 + topRows, meter + vorlauf + 49))
                .Interior.Color = 16777215 'Hintergrund
                .Font.name = "Courier New"
                .Font.Size = 12
                .IndentLevel = 1 'Text eingerückt
                .NumberFormat = "@" 'Textformat
        End With
        
        With wks.Cells(startplaetze * 2 + 20 + starter + 4 + topRows, meter + vorlauf + 9)  'Überschrift
            .RowHeight = 20
            .Font.Size = 14
            .Font.Bold = True 'Fettschrift
            .IndentLevel = 1 'Text eingerückt
            .Value = txt(143)
        End With
    
        'Rahmen um die Anzeigetafel
        wks.Range(Cells(startplaetze * 2 + 20 + starter + 4 + topRows, meter + vorlauf + 9), _
            Cells(startplaetze * 2 + 20 + starter + collSlips.Count + 4 + topRows, meter + vorlauf + 49)).BorderAround ColorIndex:=0
        
        For i = 1 To collSlips.Count
            payout = True
            id = collSlips(i).id
            nm = collSlips(i).GamblerName
            st = collSlips(i).Stake
            od = collSlips(i).Odd * 10
            bt() = collSlips(i).bet
            
            'Analyse bet slips
            Dim payCash As String
            Dim payColor As Long
            For j = 1 To UBound(bt)
                If bt(j) <> arrayErgebnisse(j, 2) Then payout = False
            Next j
            If payout = False Then
                payCash = 0
                payColor = 16777215
            Else
                payCash = st / 10 * od
                payColor = 11854022
            End If
    
            
            wks.Cells(startplaetze * 2 + 20 + starter + 4 + i + topRows, meter + vorlauf + 9) _
                .Value = id & " - " & nm & " - " & txt(153) & ": " & Format(st, "0.00") & " " & txt(151) _
                    & " - " & txt(154) & ": " & Format(payCash, "0.00") & " " & txt(151)
            
            wks.Range(Cells(startplaetze * 2 + 20 + starter + 4 + i + topRows, meter + vorlauf + 9), _
                Cells(startplaetze * 2 + 20 + starter + 4 + i + topRows, meter + vorlauf + 49)) _
                .Interior.Color = payColor
                
        Next i
    End If
End Sub

'Formatierung für Tool-Infos
Private Sub ToolInfoFormatierung()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Formatierung für Tool-Infos
        wks.Range(Columns(meter + vorlauf + 45), Columns(meter + vorlauf + 150)).ColumnWidth = 2
        With wks.Range(Cells(startplaetze * 2 + 10 + starter + 136 + topRows, meter + vorlauf + 65), _
            Cells(startplaetze * 2 + 10 + starter + 210 + topRows, meter + vorlauf + 150))
            .Font.name = "Courier New"
            .Font.Size = 12
            .EntireRow.RowHeight = 12
'            .Value = "o" 'für Tests während der Entwicklung einkommentieren
        End With
            
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Tool-Infos anzeigen
Private Sub ToolInfo()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Draw logo (get data from worksheet "Pic")
        k = 2
        For i = 1 To 21
            For j = 1 To 41
                wks.Cells(startplaetze * 2 + starter + 146 + i + topRows, meter + vorlauf + 64 + j).Interior.Color _
                    = wksPic.Cells(k, 2).Value
                k = k + 1
            Next j
        Next i
    
    'Info-Text
    For i = 1 To 5
        wks.Cells(startplaetze * 2 + starter + 162 + i + topRows, meter + vorlauf + 108).Value = txt(32 + i)
    Next i
    
    'Info-Text (technische Details)
    For i = 1 To 50
        wks.Cells(startplaetze * 2 + starter + 169 + i + topRows, meter + vorlauf + 65).Value = txt(38 + i)
    Next i
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Warnhinweis
Private Sub Warnhinweis()
    Dim warn As String
    
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    warn = warn & txt(30) & vbNewLine & txt(31)
    MsgBox warn, vbExclamation, tool

    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Beenden
Private Sub Schliessen()
    On Error Resume Next
        'Tabellenblatt wieder löschen
        Application.DisplayAlerts = False 'Warnmeldungen ausschalten
        wks.Delete 'Tabellenblatt löschen
        Application.DisplayAlerts = True 'Warnmeldungen einschalten
    On Error GoTo 0
End Sub

'Information for UserForm "Start"
Private Sub UserFormSTART()
    With frmStart
        .lblS1.Caption = race & " " & jahr
        .lblS2.Caption = raceType & " " & txt(1) & " " & meter & " " & txt(3)
        .lblS3.Caption = rennbahn & " " & txt(10) & " " & rennort
        .lblS4.Caption = gelaeuf
        .lblS4.BackColor = trackCol
        .lblS6.Caption = starter & " " & txt(100)
        .lblS8.Caption = txt(104)
        .btnS1.Caption = txt(101)
        .btnS2.Caption = txt(102)
        .btnS6.Caption = txt(107)
        .height = 327
        .Show 'Display UserForm
    End With
End Sub

'Name of the gambler for placing a bet
Public Sub Gambler()
    Dim strGambler As String
    strGambler = InputBox(txt(111), race & " " & jahr)
    If Trim(strGambler) <> "" Then Call UserFormBetSlip(strGambler) 'if name is not empty
End Sub

'Information for UserForm "BetSlip"
Private Sub UserFormBetSlip(g As String)
    With frmBetSlip
        .Caption = g
        .lblC1 = rennbahn & " - " & rennort
        .lblC2 = race & " " & jahr
        .Show 'Display UserForm
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
        .Caption = race & " " & jahr
        .height = 60
    End With
    
    k = 1
    
    With frmOdds
        .lblO0.Caption = txt(160)
        .lblO1.Caption = txt(161)
        .lblO2.Caption = txt(162)
        .lblO3.Caption = txt(163)
    End With

    For i = 1 To UBound(arrayPferde)
        If min = 0 Or arrayPferde(i, 17) < min Then min = arrayPferde(i, 17)
        If arrayPferde(i, 17) > max Then max = arrayPferde(i, 17)
    Next i
    
    For i = min To max
        For j = 1 To UBound(arrayPferde)
            If arrayPferde(j, 17) = i Then
            
                Set objLabel1 = frmOdds.Controls.Add("Forms.Label.1", , True)
                With objLabel1 'Nr and name of the horse
                    .Font.name = "Tahoma"
                    .Font.Size = 12
                    .left = 12
                    .top = 12 + k * 18
                    .width = 200
                    .TextAlign = fmTextAlignLeft
                    .Caption = "#" & arrayPferde(j, 11) & vbTab & arrayPferde(j, 1)
                    If arrayPferde(j, 0) <> "START" Then .Font.Strikethrough = True
                End With
                
                'Adjust UserForm height
                frmOdds.height = frmOdds.height + objLabel1.height

                Set objLabel2 = frmOdds.Controls.Add("Forms.Label.1", , True)
                With objLabel2 'odd
                    .Font.name = "Tahoma"
                    .Font.Size = 12
                    .left = 220
                    .top = 12 + k * 18
                    .width = 62
                    .TextAlign = fmTextAlignRight
                    .Caption = arrayPferde(j, 17) & ":10"
                    If arrayPferde(j, 0) <> "START" Then .Font.Strikethrough = True
                End With
                
                If arrayPferde(j, 0) = "START" Then
                    Set objLabel3 = frmOdds.Controls.Add("Forms.Label.1", , True)
                    With objLabel3 '~condition
                        .Font.name = "Tahoma"
                        .Font.Size = 12
                        .left = 290
                        .top = 13 + k * 18
                        .height = 15
                        .width = 100 + (((arrayPferde(j, 6) * 1000000) - 1499900) / 2) + arrayPferde(j, 18)
                        .TextAlign = fmTextAlignLeft
                        .BackColor = 6740479
''For development purposes only
'                        .Caption = "T: " & arrayPferde(j, 6) * 1000000 & " ... " & arrayPferde(j, 18)
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
'                    .Width = 100 + (((arrayPferde(j, 6) * 1000000) - 1499900) / 2)
'                    .TextAlign = fmTextAlignLeft
'                    .BackColor = 6740479
'                    .Caption = "T: " & arrayPferde(j, 6) * 1000000
'                End With
                
                k = k + 1
            End If
        Next j
    Next i

    frmOdds.Show
End Sub

'Call UserForm for showing receipt
Public Sub Receipt(id As Integer)
    Call UserFormReceipt(id)
End Sub

'Information for UserForm "Receipt"
Private Sub UserFormReceipt(id As Integer)
    Dim bet As String, horsename As String
    Dim i As Integer, j As Integer
    Dim bt() As Integer
    bt() = collSlips(id).bet
    For i = 1 To UBound(bt)
        For j = 1 To UBound(arrayPferde)
            If arrayPferde(j, 11) = bt(i) Then
                horsename = arrayPferde(j, 1)
                Exit For
            End If
        Next j
        bet = bet & bt(i) & " " & horsename & vbNewLine
    Next i

    With frmReceipt
        .Caption = collSlips(id).GamblerName
        .lblR1 = UCase(rennort)
        .lblR2 = UCase(race)
        .lblR3 = starter & " " & UCase(txt(100))
        .lblR4 = UCase(collSlips(id).BType)
        .lblR5 = UCase(bet)
        .lblR6 = UCase(txt(152) & " " & Format(collSlips(id).Stake, "0.00") & " " & txt(151))
        .lblR7 = collSlips(id).id
        .Show 'Display UserForm
    End With
End Sub

'Globale Fehlerbehandlung
Private Sub Programmabsturz()
    MsgBox txt(99), , tool
End Sub


'CALLBACKS (Aktionen im Menüband)
'--------------------------------

'Callback for customUI.onLoad
Public Sub AI_GALOPPSIMAddinInitialize(ribbon As IRibbonUI)
    Set GaloppSimRibbon = ribbon
    If wksTxt Is Nothing Then
        Set wksTxt = ThisWorkbook.Worksheets("Txt")
        Call Texte
    End If
End Sub

'Callback for lblDE getLabel
Public Sub AI_GetLabel(control As IRibbonControl, ByRef returnedVal)
    Select Case control.id
        Case "btn8GALOPPSIM" 'button "deutsch"
            returnedVal = txt(230)
        Case "btn9GALOPPSIM" 'button "english"
            returnedVal = txt(231)
        Case "btn10GALOPPSIM" 'button "russian"
            returnedVal = txt(232)
        Case "group1GALOPPSIM" 'group "Settings"
            returnedVal = txt(361)
        Case "group5GALOPPSIM" 'group "Horse race"
            returnedVal = txt(362)
        Case "group2GALOPPSIM" 'group "After the race"
            returnedVal = txt(360)
        Case "menu1GALOPPSIM" 'menu "Racing options"
            returnedVal = txt(252)
        Case "menu2GALOPPSIM" 'menu "Excel options"
            returnedVal = txt(253)
        Case "menu3GALOPPSIM" 'menu "Language"
            returnedVal = txt(229)
        Case "chk1GALOPPSIM" 'checkbox "Racing tactics"
            returnedVal = txt(194)
        Case "chk3GALOPPSIM" 'checkbox "Hoofprints"
            returnedVal = txt(248)
        Case "chk6GALOPPSIM" 'checkbox "Zoom"
            returnedVal = txt(249)
        Case "chk2GALOPPSIM" 'checkbox "Namen der Pferde im Rennen anzeigen"
            returnedVal = txt(256)
        Case "chk10GALOPPSIM" 'checkbox "Farben der Pferde im Rennen darstellen"
            returnedVal = txt(257)
        Case "chk11GALOPPSIM" 'checkbox "Favorit im Rennen markieren"
            returnedVal = txt(258)
        Case "chk8GALOPPSIM" 'checkbox "Namen der Pferde im Ziel anzeigen"
            returnedVal = txt(259)
        Case "chk12GALOPPSIM" 'checkbox "Farben im Klassement darstellen"
            returnedVal = txt(266)
        Case "chk4GALOPPSIM" 'checkbox "Anzeige des Klassements schrittweise"
            returnedVal = txt(267)

        Case "chk5GALOPPSIM" 'checkbox "Maximize window"
            returnedVal = txt(271)
        Case "chk9GALOPPSIM" 'checkbox "Hide row and column numbers"
            returnedVal = txt(272)
        Case "chk7GALOPPSIM" 'checkbox "Hide Excel gridlines"
            returnedVal = txt(273)

        Case "btn1GALOPPSIM" 'button "Start"
            returnedVal = txt(278)
        Case "comboInstRennen" 'combobox "Race selection"
            returnedVal = txt(288)
        Case "btn2GALOPPSIM" 'button "Ranking list"
            returnedVal = txt(279)
        Case "btn3GALOPPSIM" 'button "Fotofinish"
            returnedVal = txt(280)
        Case "btn7GALOPPSIM" 'button "Betting analysis"
            returnedVal = txt(281)
        Case "btn5GALOPPSIM" 'button "Info"
            returnedVal = txt(282)
        Case "btn6GALOPPSIM" 'button "Warning"
            returnedVal = txt(283)
        Case "btn11GALOPPSIM" 'button "Movie"
            returnedVal = txt(284)
        Case "btn4GALOPPSIM" 'button "Close"
            returnedVal = txt(285)
    End Select
End Sub

'Callback for menuCaptionGALOPPSIM2 getTitle
Public Sub AI_GetTitle(control As IRibbonControl, ByRef returnedVal)

    Select Case control.id
        Case "menuCaptionGALOPPSIM2"
            returnedVal = txt(292)
        Case "menuCaptionGALOPPSIM3"
            returnedVal = txt(293)
        Case "menuCaptionGALOPPSIM4"
            returnedVal = txt(294)
    End Select
    
End Sub

'Callbacks for Tooltips
Public Sub AI_GetScreentip(control As IRibbonControl, ByRef screentip)
   
    Select Case control.id
    
        Case "chk1GALOPPSIM"
            screentip = txt(300) 'racing tactics
        Case "chk3GALOPPSIM"
            screentip = txt(302) 'hoofprints
        Case "chk6GALOPPSIM"
            screentip = txt(304) 'zoom
        Case "chk2GALOPPSIM"
            screentip = txt(306) 'horse names in the race
        Case "chk10GALOPPSIM"
            screentip = txt(308) 'horse colours in the race
        Case "chk11GALOPPSIM"
            screentip = txt(310) 'highlight favourite
        Case "chk8GALOPPSIM"
            screentip = txt(312) 'horse names in finish
        Case "chk12GALOPPSIM"
            screentip = txt(314) 'colours in ranking list
        Case "chk4GALOPPSIM"
            screentip = txt(316) 'delay ranking list
        Case "chk5GALOPPSIM"
            screentip = txt(318) 'maximize window
        Case "chk9GALOPPSIM"
            screentip = txt(320) 'hide rox/col numbers
        Case "chk7GALOPPSIM"
            screentip = txt(322) 'hide excel grid
        Case "btn1GALOPPSIM"
            screentip = txt(324) 'start button
        Case "comboInstRennen"
            screentip = txt(326) 'race selection combobox
        Case "btn2GALOPPSIM"
            screentip = txt(328) 'ranking list button
        Case "btn3GALOPPSIM"
            screentip = txt(330) 'fotofinish button
        Case "btn7GALOPPSIM"
            screentip = txt(332) 'betting analysis button
        Case "btn5GALOPPSIM"
            screentip = txt(334) 'info button
        Case "btn6GALOPPSIM"
            screentip = txt(336) 'warning button
        Case "btn11GALOPPSIM"
            screentip = txt(338) 'movie button
        Case "btn4GALOPPSIM"
            screentip = txt(340) 'close button
    End Select

End Sub

Public Sub AI_GetSupertip(control As IRibbonControl, ByRef supertip)
   
    Select Case control.id
    
        Case "chk1GALOPPSIM"
            supertip = txt(301) 'racing tactics
        Case "chk3GALOPPSIM"
            supertip = txt(303) 'hoofprints
        Case "chk6GALOPPSIM"
            supertip = txt(305) 'zoom
        Case "chk2GALOPPSIM"
            supertip = txt(307) 'horse names in the race
        Case "chk10GALOPPSIM"
            supertip = txt(309) 'horse colours in the race
        Case "chk11GALOPPSIM"
            supertip = txt(311) 'highlight favourite
        Case "chk8GALOPPSIM"
            supertip = txt(313) 'horse names in finish
        Case "chk12GALOPPSIM"
            supertip = txt(315) 'colours in ranking list
        Case "chk4GALOPPSIM"
            supertip = txt(317) 'delay ranking list
        Case "chk5GALOPPSIM"
            supertip = txt(319) 'maximize window
        Case "chk9GALOPPSIM"
            supertip = txt(321) 'hide rox/col numbers
        Case "chk7GALOPPSIM"
            supertip = txt(323) 'hide excel grid
        Case "btn1GALOPPSIM"
            supertip = txt(325) 'start button
        Case "comboInstRennen"
            supertip = txt(327) 'race selection combobox
        Case "btn2GALOPPSIM"
            supertip = txt(329) 'ranking list button
        Case "btn3GALOPPSIM"
            supertip = txt(331) 'fotofinish button
        Case "btn7GALOPPSIM"
            supertip = txt(333) 'betting analysis button
        Case "btn5GALOPPSIM"
            supertip = txt(335) 'info button
        Case "btn6GALOPPSIM"
            supertip = txt(337) 'warning button
        Case "btn11GALOPPSIM"
            supertip = txt(339) 'movie button
        Case "btn4GALOPPSIM"
            supertip = txt(341) 'close button
    End Select

End Sub

'Startbutton im Menüband
Public Sub AI_START(control As IRibbonControl)
    'Leaving the current race?
    If rennen Then
        Dim r As String
        With ThisWorkbook.Worksheets(Auswahl)
            r = .Cells(2, 2).Value & " " & _
                .Cells(3, 2).Value & " (" & _
                .Cells(7, 2).Value & "m)"
        End With
        If MsgBox((txt(11) & " " & r), vbOKCancel, tool) = vbCancel Then Exit Sub
    End If

    Call derby
End Sub

'Ergebnis-Button im Menüband
Public Sub AI_ERGEBNIS(control As IRibbonControl)
    Call ergebnis_show
End Sub

Private Sub ergebnis_show()
    If rennen Then
        Call ScrollErgebnis(startplaetze * 2 + 9)
        If PlayMode = "RS" Then frmRSnavi.Show
    Else
        Call Grunddaten2
        Call Texte 'Texte einlesen
        MsgBox txt(156), , race & " " & jahr
    End If
End Sub

'Zielfoto-Button im Menüband
Public Sub AI_FOTO(control As IRibbonControl)
    Call foto_show
End Sub

Private Sub foto_show()
    If rennen Then
        'Text anpassen
            wks.Cells(4 + topRows, meter + vorlauf + 9).Value = txt(28) '("Zielfoto")
        Call ScrollErgebnis(1)
        Call FotoZeigen
        If PlayMode = "RS" Then frmRSnavi.Show
    Else
        Call Grunddaten2
        Call Texte 'Texte einlesen
        MsgBox txt(156), , race & " " & jahr
    End If
End Sub

'Wett-Button im Menüband
Public Sub AI_WETTEN(control As IRibbonControl)
    Call wetten_show
End Sub

Private Sub wetten_show()
    If rennen And betting Then
        Call ScrollErgebnis(startplaetze * 2 + 19 + starter + 4)
        If PlayMode = "RS" Then frmRSnavi.Show
    Else
        Call Grunddaten2
        Call Texte 'Texte einlesen
        MsgBox txt(155), , race & " " & jahr
    End If
End Sub

'Info-Button im Menüband
Public Sub AI_INFO(control As IRibbonControl)
    Call info_show
End Sub

Private Sub info_show()
    If PlayMode = "AI" Then
        If Not rennen And Not infoblatt Then
            Call Grunddaten1
            Call Grunddaten2
            Call Texte
            infoblatt = True
        End If
    End If
    
    Call ToolInfoFormatierung
    Call ScrollInfo
    Call ToolInfo
    
    If PlayMode = "RS" Then frmRSnavi.Show
End Sub

'Warnhinweis-Button im Menüband
Public Sub AI_WARNUNG(control As IRibbonControl)
    Call warning_show
End Sub

Private Sub warning_show()
    Call Grunddaten2
    Call Texte
    Call Warnhinweis
End Sub

'Movie button (Ribbon)
Public Sub AI_MOVIE1(control As IRibbonControl)
    Call movie1
End Sub

Private Sub movie1()
    Call modMovie1.movie1(txt(), ActiveSheet.name)
End Sub

'Ende-Button im Menüband
Public Sub AI_ENDE(control As IRibbonControl)
    If rennen Or infoblatt Then
        'Variablen zurücksetzen
            rennen = False
            infoblatt = False
        'Blatt löschen
            Call Schliessen
    End If
End Sub

'Language button "DE"
Public Sub AI_DE(control As IRibbonControl)
    Call languageDE
End Sub

Private Sub languageDE()
    language = "DE"
    Call changeLanguage
End Sub

'Language button "EN"
Public Sub AI_EN(control As IRibbonControl)
    Call languageEN
End Sub

Private Sub languageEN()
    language = "EN"
    Call changeLanguage
End Sub

'Language button "RU"
Public Sub AI_RU(control As IRibbonControl)
    Call languageRU
End Sub

Private Sub languageRU()
    language = "RU"
    Call changeLanguage
End Sub

Private Sub changeLanguage()
    Dim coll As OLEObject
    
    Call Texte 'Get new texts
    
    If PlayMode = "RS" Then
        Call RS_headings
        
        For Each coll In wks.OLEObjects
            If coll.name <> "CBraces" Then 'no action at dropdown with the races
                Call newText(coll.name)
            End If
        Next coll
    Else ' AI mode
        GaloppSimRibbon.Invalidate
    End If

End Sub

Private Sub newText(name As String)
    Dim txTr As Integer, txFa As Integer
    
    Select Case name

    'Main buttons
        Case "derby"
            wks.OLEObjects(name).Object.Caption = txt(184)

        Case "results"
            wks.OLEObjects(name).Object.Caption = txt(185)

        Case "fotofinish"
            wks.OLEObjects(name).Object.Caption = txt(186)

        Case "bets"
            wks.OLEObjects(name).Object.Caption = txt(187)

        Case "info"
            wks.OLEObjects(name).Object.Caption = txt(188)

        Case "warning"
            wks.OLEObjects(name).Object.Caption = txt(189)

        Case "movie1"
            wks.OLEObjects(name).Object.Caption = txt(220)
            
    'Languages
        Case "DE"
            wks.OLEObjects(name).Object.Caption = txt(230)
        Case "EN"
            wks.OLEObjects(name).Object.Caption = txt(231)
        Case "RU"
            wks.OLEObjects(name).Object.Caption = txt(232)
        
    'Option buttons
        Case "bet"
            '... = name, state true/false, caption if false, caption if true
            Call newTextOptionButtons(name, betting, 203, 193)
            
        Case "tac"
            Call newTextOptionButtons(name, taktik, 204, 194)
            
        Case "huf"
            Call newTextOptionButtons(name, hufspur, 205, 195)

        Case "nameS"
            Call newTextOptionButtons(name, namen, 206, 196)
            
        Case "colS"
            Call newTextOptionButtons(name, pferdefarbe, 207, 197)
            
        Case "mFav"
            Call newTextOptionButtons(name, markFav, 208, 198)
            
        Case "nameF"
            Call newTextOptionButtons(name, namen2, 209, 199)
            
        Case "colR"
            Call newTextOptionButtons(name, colResult, 210, 200)
            
        Case "vzRes"
            Call newTextOptionButtons(name, spannung, 211, 201)
            
        Case "zoom"
            Call newTextOptionButtons(name, zoom, 212, 202)
            
        Case "maxi"
            Call newTextOptionButtons(name, maxi, 213, 190)
            
        Case "RCHide"
            Call newTextOptionButtons(name, headings, 214, 191)
            
        Case "GridHide"
            Call newTextOptionButtons(name, gitter, 215, 192)
            
    End Select
End Sub

Private Sub newTextOptionButtons(name As String, state As Boolean, txFa As Integer, txTr As Integer)
    Dim textC As String
    
    If state = True Then
        textC = txt(txTr)
    Else
        textC = txt(txFa)
    End If

    wks.OLEObjects(name).Object.Caption = textC
End Sub

'Checkboxen im Menüband (onAction)
Public Sub AI_OPTIONEN(control As IRibbonControl, pressed As Boolean)
    Select Case control.id
        Case "chk1GALOPPSIM" 'Renntaktik
            If pressed Then
                taktik = True
            Else
                taktik = False
            End If
        Case "chk2GALOPPSIM" 'Namen der Pferde am linken Rand
            If pressed Then
                namen = True
            Else
                namen = False
            End If
        Case "chk10GALOPPSIM" 'Farben der Pferde am linken Rand
            If pressed Then
                pferdefarbe = True
            Else
                pferdefarbe = False
            End If
        Case "chk8GALOPPSIM" 'Namen der Pferde im Ziel
            If pressed Then
                namen2 = True
            Else
                namen2 = False
            End If
        Case "chk3GALOPPSIM" 'Hufspuren
            If pressed Then
                hufspur = True
            Else
                hufspur = False
            End If
        Case "chk4GALOPPSIM" 'Ergebnisliste langsam aufbauen (Spannung)
            If pressed Then
                spannung = True
            Else
                spannung = False
            End If
        Case "chk5GALOPPSIM" 'Excel-Fenster maximieren
            If pressed Then
                maxi = True
            Else
                maxi = False
            End If
        Case "chk6GALOPPSIM" 'Rennbahn größer darstellen wenn Fenster groß
            If pressed Then
                zoom = True
            Else
                zoom = False
            End If
        Case "chk7GALOPPSIM" 'Gitternetzlinien auf Tabellenblatt ausblenden
            If pressed Then
                gitter = True
            Else
                gitter = False
            End If
        Case "chk9GALOPPSIM" 'Zeilen-/Spaltennummern auf Tabellenblatt ausblenden
            If pressed Then
                headings = True
            Else
                headings = False
            End If
        Case "chk11GALOPPSIM" 'Favorit rot markieren
            If pressed Then
                markFav = True
            Else
                markFav = False
            End If
        Case "chk12GALOPPSIM" 'Farben der Pferde im Klassement anzeigen"
            If pressed Then
                colResult = True
            Else
                colResult = False
            End If
    End Select
End Sub

'Initialwerte der Checkboxen im Menüband (getPressed)
Public Sub AI_OPTIONEN_INI(control As IRibbonControl, ByRef standardwert)
    Select Case control.id
        Case "chk1GALOPPSIM" 'Renntaktik
            standardwert = taktik
        Case "chk2GALOPPSIM" 'Namen der Pferde am linken Rand
            standardwert = namen
        Case "chk10GALOPPSIM" 'Farben der Pferde am linken Rand
            standardwert = pferdefarbe
        Case "chk8GALOPPSIM" 'Namen der Pferde im Ziel
            standardwert = namen2
        Case "chk3GALOPPSIM" 'Hufspuren
            standardwert = hufspur
        Case "chk4GALOPPSIM" 'Ergebnisliste langsam aufbauen
            standardwert = spannung
        Case "chk5GALOPPSIM" 'Excel-Fenster maximieren
            standardwert = maxi
        Case "chk6GALOPPSIM" 'Rennbahn größer darstellen wenn Fenster groß
            standardwert = zoom
        Case "chk7GALOPPSIM" 'Gitternetzlinien auf Tabellenblatt ausblenden
            standardwert = gitter
        Case "chk9GALOPPSIM" 'Zeilen-/Spaltennummern auf Tabellenblatt ausblenden
            standardwert = headings
        Case "chk11GALOPPSIM" 'Favorit rot markieren
            standardwert = markFav
        Case "chk12GALOPPSIM" 'Farben der Pferde im Klassement anzeigen
            standardwert = colResult
    End Select
End Sub

'Anzahl der installierten Rennen auslesen
Public Sub AI_InstRennen_getItemCount(control As IRibbonControl, ByRef returnedVal)
    
    Dim wksR As Worksheet
    Dim cnt As Long
    
    Set collRennen = Nothing
    Set collRennen = New Collection
    
    For Each wksR In ThisWorkbook.Worksheets
        If left(wksR.name, 5) = "race_" Then cnt = cnt + 1
    Next wksR
    
    returnedVal = cnt
    
End Sub

'Namen der installierten Rennen auslesen
Public Sub AI_InstRennen_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)

    Dim wksR As Worksheet
    Dim cnt As Long
    
    For Each wksR In ThisWorkbook.Worksheets
        If left(wksR.name, 5) = "race_" Then cnt = cnt + 1
        If cnt = index + 1 Then
            With wksR
                collRennen.Add .name
                returnedVal = .Cells(2, 2).Value & " " & _
                                .Cells(3, 2).Value & " (" & _
                                .Cells(7, 2).Value & "m)"
                End With
            Exit For
        End If
    Next wksR

End Sub

'Ausgewähltes Rennen setzen
Public Sub AI_InstRennen_Click(control As IRibbonControl, id As String, index As Integer)

    Auswahl = collRennen(index + 1)

End Sub

'Do not save the workbook
'https://support.microsoft.com/de-de/help/213428/how-to-suppress-save-changes-prompt-when-you-close-a-workbook-in-excel
Sub Auto_Close()
    If PlayMode = "RS" Then ThisWorkbook.Saved = True
End Sub
