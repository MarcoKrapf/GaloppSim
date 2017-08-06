Attribute VB_Name = "modDerby"
Option Explicit
Option Private Module

'Galopp-Simulator (Version 148.01GIT) - August 2017
'Marco Krapf - excel@marco-krapf.de - https://marco-krapf.de/excel/
'GNU General Public License v3.0


'Workbookweit gültige Variablen für Menüband-Elemente
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
Public Auswahl As String 'Ausgewähltes Rennen

'Modulweit gültige Konstanten und Variablen
Const tool As String = "Galopp-Simulator (Version 148.01GIT)"
Dim wks As Worksheet 'Tabellenblatt für das Rennen
Dim wksData As Worksheet 'Tabellenblatt mit den Renndaten
Dim wksTxt As Worksheet 'Tabellenblatt mit Texten
Dim wksTech As Worksheet 'Tabellenblatt mit technischen Daten
Dim wksCheck As Worksheet 'Variable zum Prüfen ob das Tabellenblatt GALOPPSIM da ist
Dim arrayPferde() As Variant 'Alle Infos über jedes Pferd
Dim arrayFotofinish() As Variant 'Position jedes Pferds wenn das erste im Ziel ankommt
Dim arrayBerechnung() As Variant 'Berechnung der Platzierung beim (gleichzeitigen) Zieleinlauf
Dim arrayErgebnisse() As Variant 'Ergebnisliste
Dim collRennen As Collection 'Installierte Rennen
Dim TEST As Boolean 'Kennzeichen ob das Programm im Testmodus während der Entwicklung läuft
Dim infoblatt As Boolean 'Kennzeichen ob die Info angezeigt wurde
Dim Rennen As Boolean 'Kennzeichen ob ein Rennen gestartet wurde
Dim race As String 'Name des Rennens
Dim rennort As String 'Ort der Rennbahn
Dim rennbahn As String 'Name der Rennbahn
Dim jahr As String 'Jahr des Rennens
Dim gelaeuf As String 'Art der Rennbahn (Grasbahn, Sandbahn, Schnee)
Dim disziplin As String 'Flachrennen oder Hindernisrennen
Dim vorlauf As Integer 'Anzahl Spalten vor den Startboxen
Dim meter As Integer 'Länge der Rennbahn in Metern
Dim spur As Integer 'Spurbreite (in Abhängigkeit der Fenstergröße)
Dim bahn As Double 'Länge der Bahn (in Abhängigkeit der Fenstergröße)
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
Dim i As Integer, j As Integer, k As Integer
Dim txt(1 To 99) As String 'Array für Textbausteine

'Tests während der Entwicklung
Public Sub TEST_Rennen()
    TEST = True 'für realistischen Test auf False stellen, bei FALSE werden die Verzögerungen übersprungen
    'Parameter, die normal aus dem Menüband ausgelesen werden
        Call Standardparameter
    'Rennen auswählen
        Auswahl = "DERBY17HH"
    'Rennen starten
        Call Derby
End Sub

Public Sub TEST_Info()
    TEST = True
    Call Grunddaten1
    Call Grunddaten2
    Call Grunddaten3
    Call Texte
    Call ToolInfoFormatierung
    Call ScrollInfo
    Call ToolInfo
End Sub

Public Sub TEST_Warnhinweis()
    TEST = True
    Call Grunddaten2
    Call Texte
    Call Warnhinweis
End Sub

'Startprozedur
Private Sub Derby()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
        Call Renndaten 'Tabellenblatt mit ausgewähltem Rennen
        Call Grunddaten2 'Tabellenblatt "Txt"
        Call Texte 'Texte einlesen
    If MsgBox("Nächstes Rennen: " & race & " " & jahr & vbNewLine & vbNewLine & _
            disziplin & " " & txt(1) & " " & meter & " " & txt(3) & vbNewLine & _
            rennbahn & " " & txt(8) & " " & rennort & " - " & gelaeuf, _
            vbOKCancel, tool) = vbOK Then
        Call Grunddaten1 'Tabellenblatt "GALOPPSIM"
        Call Grunddaten3 'Tabellenblatt "TechData"
        Call Grunddaten4 'Grundsätzliche Daten auslesen bzw. festlegen
        Call Pferdedaten 'Daten über das Rennen einlesen
        Call Rennstrecke 'Geläuf zeichnen
        Call Startfeld 'Pferde generieren
        Call Begruessung 'Popup zu Rennbeginn
        Call Startaufstellung 'Pferde in Boxen stellen
        Call Vorstellung 'Pferde vorstellen
        Call Rennstart 'Rennstart
        Call Ergebnisse 'Ergebnistafel
        Call Siegerpferd 'Grafik
    End If

    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Standardparameter setzen für Tests während der Entwicklung
Private Sub Standardparameter()
    taktik = False
    namen = True
    namen2 = True
    pferdefarbe = True
    hufspur = False
    spannung = True
    maxi = True
    gitter = False
    headings = False
    zoom = True
End Sub

'Grunddaten einlesen
Private Sub Grunddaten1()
    'Fehlerbehandlung
    On Error GoTo FEHLER

    'Prüfen ob es schon ein Tabellenblatt gibt
    For Each wksCheck In ActiveWorkbook.Worksheets
        If wksCheck.Name = "GALOPPSIM" Then 'Tabellenblatt ist schon da
            Application.DisplayAlerts = False 'Warnmeldungen ausschalten
            wksCheck.Delete 'Tabellenblatt löschen
            Application.DisplayAlerts = True 'Warnmeldungen einschalten
        End If
    Next wksCheck
    'Neues Tabellenblatt generieren
        Set wks = ActiveWorkbook.Worksheets.Add
        wks.Name = "GALOPPSIM"
        wks.Activate

    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Grunddaten einlesen
Private Sub Grunddaten2()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Tabellenblatt zuweisen
    Set wksTxt = ThisWorkbook.Worksheets("Txt")
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Grunddaten einlesen
Private Sub Grunddaten3()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Tabellenblatt zuweisen
    Set wksTech = ThisWorkbook.Worksheets("TechData")
    
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
        If zoom And ActiveWindow.Height >= 600 Then
            spur = 9 'Breite der Rennspuren
            schrift = 8 'Schriftgröße für Pferdenamen und Hufspuren
        Else
            spur = 6 'Breite der Rennspuren
            schrift = 5 'Schriftgröße für Pferdenamen und Hufspuren
        End If
    
    'Parameter in Abhängigkeit der Fensterbreite
        If zoom And ActiveWindow.Width >= 1040 Then
            bahn = 0.3 'Länge der Rennbahn
            liste = 4.5 'Breite Zielfoto/Ergebnistafel
            scrolling = 250 'Scrolling
        Else
            bahn = 0.2 'Länge der Rennbahn
            liste = 3 'Breite Zielfoto/Ergebnistafel
            scrolling = 450 'Scrolling
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
        vorlauf = 10
        
    'Variablen setzen, dass Rennen gestartet wurde
        Rennen = True
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
        race = .Cells(1, 1).Value
        rennort = .Cells(3, 1).Value
        rennbahn = .Cells(4, 1).Value
        jahr = .Cells(2, 1).Value
        gelaeuf = .Cells(5, 1).Value
        meter = .Cells(6, 1).Value
        disziplin = .Cells(7, 1).Value
    End With
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Texte einlesen
Private Sub Texte()
    For i = 1 To UBound(txt)
        txt(i) = wksTxt.Cells(i, 2).Value
    Next i
End Sub

'Daten über die Pferde
Private Sub Pferdedaten()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Anzahl der gemeldeten Pferde aus Tabellenblatt auslesen
        startplaetze = Application.WorksheetFunction.CountA(wksData.Columns(5)) - 1
    'Anzahl der Starter aus Tabellenblatt auslesen
        starter = Application.WorksheetFunction.CountIf(wksData.Columns(5), "START")
    'Arrays anlegen
    ReDim arrayPferde(1 To startplaetze, 0 To 16) 'Alle Daten der Pferde
    ReDim arrayFotofinish(1 To startplaetze, 0 To 1) 'Snapshot für Zielfoto/Fotofinish
    ReDim arrayErgebnisse(0 To starter, 0 To 5) 'Ergebnisliste
        
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

Private Sub Rennstrecke()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Bildschirmaktualisierung ausschalten
        Application.ScreenUpdating = False
    'Spalten A-D des Fensters fixieren wenn Checkbox für Namen oder Farbe der Pferde am linken Rand angehakt ist
        If namen Or pferdefarbe Then
            With ActiveWindow
                .SplitColumn = 4
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
        wks.Range(Rows(1), Rows(5)).EntireRow.RowHeight = 15 'oberhalb der Rennbahn
        wks.Range(Rows(startplaetze * 2 + 6 + 1), _
            Rows(startplaetze * 2 + 42)).EntireRow.RowHeight = 15 'unterhalb der Rennbahn
        wks.Rows(startplaetze * 2 + 10).RowHeight = 20 'Überschrift Ergebnistafel
    'Renndaten anzeigen
        With wks.Cells(2, 5)
            .Font.Name = "Arial Black"
            .IndentLevel = 1 'Text eingerückt
            .Value = race & " " & jahr & " - " & rennbahn & ", " & rennort
        End With
        With wks.Cells(3, 5)
            .Font.Name = "Arial"
            .Font.Bold = True 'Fettschrift
            .IndentLevel = 1 'Text eingerückt
            .Value = disziplin & " " & txt(1) & " " & _
                meter & txt(2) & " - " & gelaeuf
        End With
    'Bereich vor Startboxen
        wks.Columns(1).ColumnWidth = 1 'Linker Rand
        wks.Columns(2).ColumnWidth = 1.5 'Spalte für Farbe der Pferde
        wks.Columns(3).ColumnWidth = 4.5 'Spalte für Startnummern
        wks.Columns(4).ColumnWidth = 14 'Spalte für Pferdenamen
        wks.Range(Columns(5), Columns(vorlauf + 5)).ColumnWidth = bahn
        wks.Columns(vorlauf - 3).ColumnWidth = 6 'Spalte für Boxennummern
    'Rennbahn
        wks.Range(Columns(vorlauf + 6), Columns(meter + vorlauf + 6)).ColumnWidth = bahn 'Länge der Bahn
        wks.Range(Rows(6), Rows(startplaetze * 2 + 6)).RowHeight = spur 'Spurbreite
        wks.Range(Cells(5, 1), Cells(startplaetze * 2 + 7, meter + vorlauf + 7)).Interior.ColorIndex = 43 'grün
    'Schriftart
        wks.Range(Cells(4, 1), Cells(startplaetze * 2 + 8, meter + vorlauf + 7)).Font.Name = "Arial"
    'Startboxen
        For i = 6 To (startplaetze * 2 + 6) Step 2 'Für jeden Startplatz eine Box
            wks.Range(Cells(i, vorlauf + 1), Cells(i, vorlauf + 6)).Interior.ColorIndex = 1 'schwarz
        Next i
        wks.Range(Cells(6, vorlauf + 6), Cells(startplaetze * 2 + 6, vorlauf + 6)).Interior.ColorIndex = 1 'schwarz
    'Beschriftung der Boxen
        With wks.Range(Cells(7, vorlauf - 3), Cells(startplaetze * 2 + 5, vorlauf - 3))
            .Font.ColorIndex = 1 'schwarz
            .Font.Size = schrift
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        For i = 1 To (startplaetze) 'Für jeden Startplatz eine Box
            wks.Cells(5 + 2 * i, vorlauf - 3).Value = txt(5) & i
        Next i
    'Meter-Anzeigen
        For i = 500 To meter Step 500
'            wks.Range(Cells(5, i + vorlauf + 5), Cells(45, i + vorlauf + 5)).Interior.ColorIndex = 1 'für Entwicklung einkommentieren
            With wks.Cells(4, i + vorlauf)
                .Font.Name = "Arial"
                .Font.Bold = True 'Fettschrift
                .Value = i & txt(2) '"m"
            End With
            With wks.Cells(startplaetze * 2 + 8, i + vorlauf)
                .Font.Name = "Arial"
                .Font.Bold = True 'Fettschrift
                .Value = i & txt(2) '"m"
            End With
        Next i
    'Formatierung für Pferdenamen am linken Rand
        With wks.Range(Cells(5, 3), Cells(startplaetze * 2 + 7, 4))
            .Font.ColorIndex = 43 'grün, damit zuerst nicht sichtbar
            .IndentLevel = 1 'Text eingerückt
            .Font.Size = schrift
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    '"Schriftgröße" für Hufspuren
        With wks.Range(Cells(5, vorlauf - 1), Cells(startplaetze * 2 + 7, meter + vorlauf))
            .Font.Size = schrift
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
    'Ziel
        wks.Columns(meter + vorlauf + 5).ColumnWidth = bahn 'Ziellinie
        wks.Range(Cells(5, meter + vorlauf + 5), _
            Cells(startplaetze * 2 + 7, meter + vorlauf + 5)).Interior.ColorIndex = 56 'grau
        With wks.Cells(4, meter + vorlauf + 1)
            .Font.Name = "Arial"
            .Font.Bold = True 'Fettschrift
            .Value = meter & txt(2) 'Meter im Ziel
        End With
        With wks.Cells(startplaetze * 2 + 8, meter + vorlauf + 1)
            .Font.Name = "Arial"
            .Font.Bold = True 'Fettschrift
            .Value = meter & txt(2)  'Meter im Ziel
        End With
    'Auslauf hinter Ziel
        wks.Columns(meter + vorlauf + 7).ColumnWidth = 18
        wks.Columns(meter + vorlauf + 8).ColumnWidth = 10
    'Formatierung für Pferdenamen im Ziel
        With wks.Range(Cells(5, meter + vorlauf + 7), Cells(startplaetze * 2 + 7, meter + vorlauf + 7))
            .Font.ColorIndex = 1 'schwarz
            .IndentLevel = 1 'Text eingerückt
            .Font.Size = schrift
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
    'Formatierungen für Zielfoto
        wks.Cells(2, meter + vorlauf + 7).Font.Name = "Arial Black" '"FOTOFINISH!"
        With wks.Cells(4, meter + vorlauf + 9) '"Zielfoto..."
            .Font.Name = "Courier New"
            .Font.Size = 14
            .Font.Bold = True 'Fettschrift
        End With
    'Formatierung für Ergebnisliste/Zielfoto
        wks.Range(Columns(meter + vorlauf + 9), Columns(meter + vorlauf + 27)).ColumnWidth = liste
    'Formatierung für Siegerfoto
        wks.Range(Columns(meter + vorlauf + 27), Columns(meter + vorlauf + 44)).ColumnWidth = 2
        With wks.Cells(startplaetze * 2 + 10, meter + vorlauf + 27)
            .Font.Name = "Courier New"
            .Font.Size = 14
            .Font.Bold = True 'Fettschrift
        End With
    'Formatierung für Tool-Infos
        Call ToolInfoFormatierung
    'Cursor weit weg platzieren
        wks.Cells(100, 1).Select
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
    
    For i = 1 To startplaetze
        arrayPferde(i, 0) = wksData.Cells(1 + i, 5).Value 'Status (START, ABSAGE...)
        arrayPferde(i, 11) = wksData.Cells(1 + i, 4).Value 'Startnummer
        arrayPferde(i, 1) = wksData.Cells(1 + i, 6).Value 'Name des Pferds
        arrayPferde(i, 2) = wksData.Cells(1 + i, 7).Value 'Rennfarbe
        arrayPferde(i, 15) = wksData.Cells(1 + i, 3).Value 'Box aus der das Pferd startet
        arrayPferde(i, 3) = 5 + 2 * wksData.Cells(1 + i, 3).Value 'Zeilennummer auf der das Pferd läuft
        arrayPferde(i, 4) = vorlauf + 5 'Fixe Startposition in der Box
        arrayPferde(i, 5) = wksData.Cells(1 + i, 8).Value 'Grundgeschwindigkeit des Pferds
                                                        '(linear von 1,50010 bis 1,49990)
        'Form des Pferds durch Zufallszahl festlegen
        Randomize 'Zufallsgenerator zurücksetzen
        arrayPferde(i, 6) = (Int((speedF1 - speedF2 + 1) * Rnd + speedF2) + 100000) / 100000 'Zufallszahl
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
    'Pferdenamen am linken Rand platzieren wenn Checkbox angehakt ist
        If namen Then
            For i = 1 To startplaetze
                If arrayPferde(i, 0) = "START" Then
                    wks.Cells(arrayPferde(i, 3), 3).Value = "#" & arrayPferde(i, 11) 'Startnummer
                    wks.Cells(arrayPferde(i, 3), 4).Value = arrayPferde(i, 1) 'Name des Pferds
                End If
            Next i
        'Optimale Spaltenbreite
            wks.Range(Columns(3), Columns(4)).EntireColumn.AutoFit
        End If
    'Pferdenamen im Ziel anzeigen wenn Checkbox angehakt ist
        If namen2 Then
            For j = 1 To startplaetze
                If arrayPferde(j, 0) = "START" Then
                    wks.Cells(arrayPferde(j, 3), meter + vorlauf + 7).Value = _
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
    
    MsgBox txt(7) & " " & rennbahn & " " & txt(8) & " " & rennort & ". " & _
            txt(9) & " " & race & " " & txt(1) & " " & meter & " " & txt(3) & "." & _
            vbNewLine & vbNewLine & _
            txt(10) & " " & starter & " " & txt(11), , tool

    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Pferde in Boxen stellen
Private Sub Startaufstellung()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    If TEST = False Then Application.Wait (Now + TimeValue("0:00:02")) 'Verzögerung
    For i = 1 To startplaetze
        If arrayPferde(i, 0) = "START" Then
            wks.Range(Cells(arrayPferde(i, 3), arrayPferde(i, 4)), _
                Cells(arrayPferde(i, 3), arrayPferde(i, 4) - 7)) _
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
    
    If TEST = False Then Application.Wait (Now + TimeValue("0:00:02")) 'Verzögerung
    Application.DisplayCommentIndicator = xlCommentAndIndicator 'Kommentare anschalten
    
    For i = 1 To startplaetze
        If arrayPferde(i, 0) = "START" Then
            With wks.Cells(arrayPferde(i, 3), arrayPferde(i, 4))
                .AddComment Text:="#" & arrayPferde(i, 11) & " " & arrayPferde(i, 1)
                .Comment.Shape.TextFrame.AutoSize = True
            End With
            If i = favorit1 Then
                wks.Cells(arrayPferde(i, 3), arrayPferde(i, 4)) _
                    .Comment.Shape.Fill.ForeColor.RGB = RGB(255, 0, 0)
            End If
        End If
    Next i
    
    If TEST = False Then Application.Wait (Now + TimeValue("0:00:02")) 'Verzögerung
    'Messagebox
        MsgBox txt(13) & " " & arrayPferde(favorit1, 1) & " " & txt(14) & _
                " " & arrayPferde(favorit1, 11) & txt(15) & " " & arrayPferde(favorit1, 15) & _
                " " & txt(16) & vbNewLine & vbNewLine & _
                txt(17) & " " & arrayPferde(favorit2, 1) & " " & txt(14) & " " & _
                arrayPferde(favorit2, 11) & " " & txt(18) & " " & _
                arrayPferde(favorit2, 15) & vbNewLine & _
                txt(19) & " " & _
                arrayPferde(favorit3, 1) & " " & txt(14) & " " & _
                arrayPferde(favorit3, 11) & " " & txt(18) & " " & _
                arrayPferde(favorit3, 15) & " " & txt(20) & ".", , tool
    'Kommentare ausblenden
        Application.DisplayCommentIndicator = xlNoIndicator
    'Farben der Pferde am linken Rand platzieren wenn Checkbox angehakt ist
        If pferdefarbe Then
            For i = 1 To startplaetze
                If arrayPferde(i, 0) = "START" Then
                    wks.Cells(arrayPferde(i, 3), 2).Interior.Color = arrayPferde(i, 2) 'Farbe des Pferds
                End If
            Next i
        End If
    'Pferdenamen und Startnummern am linken Rand zeigen wenn Checkbox angehakt ist
        wks.Range(Columns(3), Columns(4)).Font.ColorIndex = 1 'schwarz
    'Verzögerung
        If TEST = False Then Application.Wait (Now + TimeValue("0:00:04"))

    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Start des Galopprennens
Private Sub Rennstart()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Boxennummern entfernen
        wks.Range(Cells(7, vorlauf - 3), Cells(5 + 2 * startplaetze, vorlauf - 3)).Value = ""
    'Verzögerung
        If TEST = False Then Application.Wait (Now + TimeValue("0:00:04"))
    'Boxen aufmachen
        wks.Range(Cells(6, vorlauf + 6), Cells(startplaetze * 2 + 6, vorlauf + 6)).Interior.ColorIndex = 43 'grün
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
                    If arrayPferde(i, 0) <> "ABSAGE" Then 'nur wenn Pferd am Start ist
                        'Pferd löschen
                            Range(Cells(arrayPferde(i, 3), arrayPferde(i, 4)), _
                                Cells(arrayPferde(i, 3), arrayPferde(i, 4) - 7)) _
                                .Interior.ColorIndex = 43 'grün
                        'Neue Position des Pferds festlegen (nur wenn Pferd noch läuft)
                            If arrayPferde(i, 0) = "START" Then
                                arrayPferde(i, 4) = arrayPferde(i, 4) + arrayPferde(i, 9)
                            End If
                        'Pferd neu setzen (auch die, die schon im Ziel sind wegen dem Rendering)
                            Range(Cells(arrayPferde(i, 3), arrayPferde(i, 4)), _
                                Cells(arrayPferde(i, 3), arrayPferde(i, 4) - 7)) _
                                .Interior.Color = arrayPferde(i, 2)
                        'Wenn Checkbox angehakt ist: Hufspur zeichnen
                            If hufspur Then Cells(arrayPferde(i, 3), arrayPferde(i, 4) - 8).Value = "."
                    End If
                    'Horizontal scrollen wenn nötig
                    If arrayPferde(i, 4) > ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.Count).Column - 50 _
                        And ActiveWindow.VisibleRange.Columns(ActiveWindow.VisibleRange.Columns.Count).Column <= meter + 11 Then
                            'Scrollen
                            ActiveWindow.ScrollColumn = ActiveWindow.VisibleRange.Column + scrolling
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
                                wks.Cells(2, meter + vorlauf + 7).Value = txt(21) '"FOTOFINISH!"
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
                If TEST = False Then Application.Wait (Now + TimeValue("0:00:04"))
            'Texte anpassen
                wks.Cells(2, meter + vorlauf + 7).Value = ""
                wks.Cells(4, meter + vorlauf + 9).Value = txt(22) '("Zielfoto wird erstellt")
            'Scrollen
                On Error Resume Next
                ActiveWindow.ScrollColumn = meter - 20
                On Error GoTo 0
            'Verzögerung
                If TEST = False Then Application.Wait (Now + TimeValue("0:00:04"))
            'Zielfoto anzeigen
                Call FotoZeigen
            'Text anpassen
                wks.Cells(4, meter + vorlauf + 9).Value = ""
                wks.Cells(4, meter + vorlauf + 9).Value = txt(23) '("Zielfoto wird ausgewertet")
            'Verzögerung
                If TEST = False Then Application.Wait (Now + TimeValue("0:00:04"))
            'Text anpassen
                wks.Cells(4, meter + vorlauf + 9).Value = txt(28) '("Zielfoto")
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
                With wks.Range(Cells(5, meter + vorlauf + 7), _
                    Cells(startplaetze * 2 + 7, meter + vorlauf + 7))
                        .Interior.ColorIndex = 1 'schwarz
                        .Interior.ColorIndex = 0 'weiß
                End With
            Next k
            wks.Range(Cells(5, meter + vorlauf + 7), _
                Cells(startplaetze * 2 + 7, meter + vorlauf + 7)).Interior.ColorIndex = 43 'grün
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
            wks.Range(Cells(5, meter + vorlauf + 9), Cells(startplaetze * 2 + 7, meter + vorlauf + 24)).Interior.ColorIndex = 1 'Rasen schwarz
            wks.Range(Cells(5, meter + vorlauf + 22), Cells(startplaetze * 2 + 7, meter + vorlauf + 22)).Interior.ColorIndex = 0 'Zielline weiß
        Else
            'Geläuf und Ziellinie (Originalfarben)
            wks.Range(Cells(5, meter + vorlauf + 9), Cells(startplaetze * 2 + 7, meter + vorlauf + 24)).Interior.ColorIndex = 43 'Rasen grün
            wks.Range(Cells(5, meter + vorlauf + 22), Cells(startplaetze * 2 + 7, meter + vorlauf + 22)).Interior.ColorIndex = 56 'Zielline grau
        End If
    'Pferde zeichnen
        Dim p As Integer 'Pferd auf Zielfoto
        For i = 1 To UBound(arrayFotofinish)
            If arrayFotofinish(i, 1) >= meter + vorlauf + 5 - 13 Then 'nur wenn Pferd im Foto ist
                If arrayFotofinish(i, 1) - 7 >= meter + vorlauf + 5 - 13 Then
                    p = arrayFotofinish(i, 1) - 7
                Else
                    p = meter + 2
                End If
                Range(Cells(arrayFotofinish(i, 0), arrayFotofinish(i, 1) + 17), _
                    Cells(arrayFotofinish(i, 0), p + 17)) _
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
        If TEST = False Then Application.Wait (Now + TimeValue("0:00:02"))

    MsgBox txt(24) & vbNewLine & txt(25), , tool
    'Fixierung des Fensters wieder aufheben wenn Checkbox für Namen oder Farbe der Pferde am linken Rand angehakt ist
        If namen Or pferdefarbe Then
            With ActiveWindow
                .SplitColumn = 0
                .SplitRow = 0
                .FreezePanes = False
            End With
        End If
    'Scrollen zu Ergebnistafel
        Call ScrollErgebnis(startplaetze * 2 + 9)
    'Anzeigetafel
        With wks.Range(Cells(startplaetze * 2 + 10, meter + vorlauf + 9), _
            Cells(startplaetze * 2 + 10 + starter + 1, meter + vorlauf + 24))
                .Interior.Color = 16777215 'Hintergrund
                .Font.Name = "Courier New"
                .Font.Size = 12
                .NumberFormat = "@" 'Textformat
        End With
        With wks.Cells(startplaetze * 2 + 10, meter + vorlauf + 9) 'Überschrift
            .Font.Size = 14
            .Font.Bold = True 'Fettschrift
            .IndentLevel = 1 'Text eingerückt
        End With

    'Rahmen um die Anzeigetafel
        wks.Range(Cells(startplaetze * 2 + 10, meter + vorlauf + 9), _
            Cells(startplaetze * 2 + 10 + starter + 1, meter + vorlauf + 24)).BorderAround ColorIndex:=0
    
    'Ergebnisse eintragen
        'Überschrift
            wks.Cells(startplaetze * 2 + 10, meter + vorlauf + 9).Value = _
                race & " (" & meter & txt(2) & ")"
        'Verzögerung
            If spannung Then Application.Wait (Now + TimeValue("0:00:04"))
        'Platzierungen anzeigen
            For i = UBound(arrayErgebnisse) To 1 Step -1
                Range(Cells(startplaetze * 2 + 10 + i, meter + vorlauf + 10), _
                    Cells(startplaetze * 2 + 10 + i, meter + vorlauf + 11)) _
                        .Interior.Color = arrayErgebnisse(i, 4) 'Farbe des Pferds
                Cells(startplaetze * 2 + 10 + i, meter + vorlauf + 13).Value = arrayErgebnisse(i, 1) & "." 'Platzierung
                Cells(startplaetze * 2 + 10 + i, meter + vorlauf + 14).Value = arrayErgebnisse(i, 3) & _
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

'Scrollen zu Zielfoto bzw. Ergebnistafel
Private Sub ScrollErgebnis(z As Integer)
    On Error Resume Next
        ActiveWindow.ScrollColumn = meter - 20
        ActiveWindow.ScrollRow = z
    On Error GoTo 0
End Sub

'Scrollen zu Tool-Info
Private Sub ScrollInfo()
    On Error Resume Next
        ActiveWindow.ScrollColumn = meter + vorlauf + 44 + 30 - 10
        ActiveWindow.ScrollRow = startplaetze * 2 + starter + 20
    On Error GoTo 0
End Sub

'Pferd mit Siegerkranz zeichnen
Private Sub Siegerpferd()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Verzögerung
        If TEST = False Then Application.Wait (Now + TimeValue("0:00:01"))
    'Bildschirmaktualisierung ausschalten
        Application.ScreenUpdating = False
    'Pferd zeichnen (auslesen aus Tabellenblatt TechData)
        For i = 1 To 13  'Zeilen
            For j = 1 To 18 'Spalten
                Select Case wksTech.Cells(i, j)
                    Case 0 'Hintergrund
                    
                    Case 1 'Kopf
                        wks.Cells(startplaetze * 2 + 11 + i, meter + vorlauf + 26 + j).Interior.Color = 36799
                    Case 2  'Mähne
                        wks.Cells(startplaetze * 2 + 11 + i, meter + vorlauf + 26 + j).Interior.Color = 24704
                    Case 3 'Siegerkranz
                        wks.Cells(startplaetze * 2 + 11 + i, meter + vorlauf + 26 + j).Interior.Color = 3506772
                    Case 4 'Auge
                        wks.Cells(startplaetze * 2 + 11 + i, meter + vorlauf + 26 + j).Interior.Color = 0
                    Case Else
                    
                End Select

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
        wks.Cells(startplaetze * 2 + 10, meter + vorlauf + 27).Value = txt(27) & " " & sieger
'        wks.Cells(startplaetze * 2 + 10, meter + vorlauf + 27).Value = txt(27) & " " & UCase(arrayErgebnisse(1, 3))
    'Bildschirmaktualisierung einschalten
        Application.ScreenUpdating = True
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Formatierung für Tool-Infos
Private Sub ToolInfoFormatierung()
    'Fehlerbehandlung
    On Error GoTo FEHLER
    
    'Formatierung für Tool-Infos
        wks.Range(Columns(meter + vorlauf + 45), Columns(meter + vorlauf + 150)).ColumnWidth = 2
        With wks.Range(Cells(startplaetze * 2 + 10 + starter + 10, meter + vorlauf + 65), _
            Cells(startplaetze * 2 + 10 + starter + 80, meter + vorlauf + 150))
            .Font.Name = "Courier New"
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
    
    'Logo zeichnen (auslesen aus Tabellenblatt TechData)
    For i = 15 To 35  'Zeilen
        For j = 1 To 41 'Spalten
            Select Case wksTech.Cells(i, j)
                Case 0 'Hintergrund
                
                Case 1 'Kopf
                    wks.Cells(startplaetze * 2 + starter + 6 + i, meter + vorlauf + 64 + j).Interior.Color = 36799
                Case 2  'Mähne
                    wks.Cells(startplaetze * 2 + starter + 6 + i, meter + vorlauf + 64 + j).Interior.Color = 24704
                Case 4 'Auge
                    wks.Cells(startplaetze * 2 + starter + 6 + i, meter + vorlauf + 64 + j).Interior.Color = 0
                Case 5 'Jockey
                    wks.Cells(startplaetze * 2 + starter + 6 + i, meter + vorlauf + 64 + j).Interior.Color = 192
                Case 6 'Jockey Kopf
                    wks.Cells(startplaetze * 2 + starter + 6 + i, meter + vorlauf + 64 + j).Interior.Color = 11389944
                Case 7 'Sattel
                    wks.Cells(startplaetze * 2 + starter + 6 + i, meter + vorlauf + 64 + j).Interior.Color = 13224393
                Case 8 'Logo-Schriftzug SIM
                    wks.Cells(startplaetze * 2 + starter + 6 + i, meter + vorlauf + 64 + j).Interior.Color = 9359529
                Case Else
                
            End Select
        Next j
    Next i
    
    'Info-Text
    For i = 1 To 5
        wks.Cells(startplaetze * 2 + starter + 36 + i, meter + vorlauf + 108).Value = txt(32 + i)
    Next i
    
    'Info-Text (technische Details)
    For i = 1 To 40
        wks.Cells(startplaetze * 2 + starter + 43 + i, meter + vorlauf + 65).Value = txt(38 + i)
    Next i
    
    Exit Sub
FEHLER:
    Call Programmabsturz
End Sub

'Warnhinweis
Private Sub Warnhinweis()
    'Fehlerbehandlung
    On Error GoTo FEHLER

    MsgBox txt(30), vbExclamation, tool

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

'Globale Fehlerbehandlung
Private Sub Programmabsturz()
    MsgBox txt(99), , tool
End Sub


'CALLBACKS (Aktionen im Menüband)
'--------------------------------

'Startbutton im Menüband
Public Sub START(control As IRibbonControl)
    Call Derby
End Sub

'Ergebnis-Button im Menüband
Public Sub ERGEBNIS(control As IRibbonControl)
    If Rennen Then Call ScrollErgebnis(startplaetze * 2 + 9)
End Sub

'Zielfoto-Button im Menüband
Public Sub FOTO(control As IRibbonControl)
    If Rennen Then
        'Text anpassen
            wks.Cells(4, meter + vorlauf + 9).Value = txt(28) '("Zielfoto")
        Call ScrollErgebnis(1)
        Call FotoZeigen
    End If
End Sub

'Info-Button im Menüband
Public Sub INFO(control As IRibbonControl)
    If Not Rennen And Not infoblatt Then
        Call Grunddaten1
        Call Grunddaten2
        Call Grunddaten3
        Call Texte
        infoblatt = True
    End If
    Call ToolInfoFormatierung
    Call ScrollInfo
    Call ToolInfo
End Sub

'Warnhinweis-Button im Menüband
Public Sub WARNUNG(control As IRibbonControl)
    Call Grunddaten2
    Call Texte
    Call Warnhinweis
End Sub

'Ende-Button im Menüband
Public Sub ENDE(control As IRibbonControl)
    If Rennen Or infoblatt Then
        'Variablen zurücksetzen
            Rennen = False
            infoblatt = False
        'Blatt löschen
            Call Schliessen
    End If
End Sub

'Checkboxen im Menüband (onAction)
Public Sub OPTIONEN(control As IRibbonControl, pressed As Boolean)
    Select Case control.id
        Case "chk1GALSIM" 'Renntaktik
            If pressed Then
                taktik = True
            Else
                taktik = False
            End If
        Case "chk2GALSIM" 'Namen der Pferde am linken Rand
            If pressed Then
                namen = True
            Else
                namen = False
            End If
        Case "chk10GALSIM" 'Farben der Pferde am linken Rand
            If pressed Then
                pferdefarbe = True
            Else
                pferdefarbe = False
            End If
        Case "chk8GALSIM" 'Namen der Pferde im Ziel
            If pressed Then
                namen2 = True
            Else
                namen2 = False
            End If
        Case "chk3GALSIM" 'Hufspuren
            If pressed Then
                hufspur = True
            Else
                hufspur = False
            End If
        Case "chk4GALSIM" 'Ergebnisliste langsam aufbauen (Spannung)
            If pressed Then
                spannung = True
            Else
                spannung = False
            End If
        Case "chk5GALSIM" 'Excel-Fenster maximieren
            If pressed Then
                maxi = True
            Else
                maxi = False
            End If
        Case "chk6GALSIM" 'Rennbahn größer darstellen wenn Fenster groß
            If pressed Then
                zoom = True
            Else
                zoom = False
            End If
        Case "chk7GALSIM" 'Gitternetzlinien auf Tabellenblatt ausblenden
            If pressed Then
                gitter = True
            Else
                gitter = False
            End If
        Case "chk9GALSIM" 'Zeilen-/Spaltennummern auf Tabellenblatt ausblenden
            If pressed Then
                headings = True
            Else
                headings = False
            End If
    End Select
End Sub

'Initialwerte der Checkboxen im Menüband (getPressed)
Public Sub OPTIONEN_INI(control As IRibbonControl, ByRef standardwert)
    Select Case control.id
        Case "chk1GALSIM" 'Renntaktik
            standardwert = taktik
        Case "chk2GALSIM" 'Namen der Pferde am linken Rand
            standardwert = namen
        Case "chk10GALSIM" 'Farben der Pferde am linken Rand
            standardwert = pferdefarbe
        Case "chk8GALSIM" 'Namen der Pferde im Ziel
            standardwert = namen2
        Case "chk3GALSIM" 'Hufspuren
            standardwert = hufspur
        Case "chk4GALSIM" 'Ergebnisliste langsam aufbauen
            standardwert = spannung
        Case "chk5GALSIM" 'Excel-Fenster maximieren
            standardwert = maxi
        Case "chk6GALSIM" 'Rennbahn größer darstellen wenn Fenster groß
            standardwert = zoom
        Case "chk7GALSIM" 'Gitternetzlinien auf Tabellenblatt ausblenden
            standardwert = gitter
        Case "chk9GALSIM" 'Zeilen-/Spaltennummern auf Tabellenblatt ausblenden
            standardwert = headings
    End Select
End Sub

'Anzahl der installierten Rennen auslesen
Public Sub InstRennen_getItemCount(control As IRibbonControl, ByRef returnedVal)
    
    Dim wksR As Worksheet
    Dim cnt As Long
    
    Set collRennen = Nothing
    Set collRennen = New Collection
    
    For Each wksR In ThisWorkbook.Worksheets
        If Left(wksR.Name, 5) = "race_" Then cnt = cnt + 1
    Next wksR
    
    returnedVal = cnt
    
End Sub

'Namen der installierten Rennen auslesen
Public Sub InstRennen_getItemLabel(control As IRibbonControl, index As Integer, ByRef returnedVal)

    Dim wksR As Worksheet
    Dim cnt As Long
    
    For Each wksR In ThisWorkbook.Worksheets
        If Left(wksR.Name, 5) = "race_" Then cnt = cnt + 1
        If cnt = index + 1 Then
            With wksR
                collRennen.Add .Name
                returnedVal = .Cells(1, 1).Value & " (" & _
                                .Cells(6, 1).Value & "m)"
                End With
            Exit For
        End If
    Next wksR

End Sub

'Ausgewähltes Rennen setzen
Public Sub InstRennen_Click(control As IRibbonControl, id As String, index As Integer)

    Auswahl = collRennen(index + 1)

End Sub
