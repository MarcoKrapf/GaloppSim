VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'VARIABLES
Private strRACE_ID As String 'Unique race ID
Private strREAL_RACE As String 'Race that took place in reality (Y/N)
Private strPARTICIPANTS As String 'Race participants (HORSE/PIG/DONKEY/DOG/UNICORN)
Private strRACE_NAME As String 'Name of the race
Private strRACE_YEAR As String 'Year of the race
Private strTRACK_LOCATION As String 'Location of the race track
Private strCOUNTRY As String 'Country of the race track
Private strTRACK_NAME As String 'Name of the race track
Private lngTRACK_COLOUR As Long 'Colour of the track surface
Private strTRACK_SURFACE As String 'Type of the race track (T = turf, D = dirt, S = snow, M = mudflats, A = artificial)
Private strTRACK_SURFACE_TEXT As String 'Full name of the surface
Private blnSQUIRT As Boolean 'Squirting water or not
Private strRACE_TYPE As String 'Race type (F = Flat, S = Steeplechase...)
Private strRACE_TYPE_TEXT As String 'Full name of the race type (Flat, Steeplechase...)
Private intMETERS As Integer 'Race distance in meters
Private strSTARTING_GATE As String 'Starting gate (Y/N)
Private intNUMBER_ENROLLED As Integer 'Number of horses registered (including horses that don´t start)
Private strLANES_FIX_OR_RANDOM As String 'Lanes fix or random (F/R)
Private strCOLOURS_FIX_OR_RANDOM As String 'Horse colours fix or random (F/R)
Private strODDS_FIX_FIX_OR_RANDOM As String 'Odds fix or random (F/R)
Private strADVERTISING As String 'Advertising (Y/N)
Private strBETTING_ALLOWED As String 'Placing bets allowed for the chosen race (Y/N)
Private strRaceSelected As String 'Selected race
Private blnRaceStarted As Boolean 'Flag indicating whether a race has been started


'PROPERTIES
'----------

Public Property Get RACE_ID() As String
    RACE_ID = strRACE_ID
End Property

Public Property Let RACE_ID(ByVal strX As String)
    strRACE_ID = strX
End Property

Public Property Get REAL_RACE() As String
    REAL_RACE = strREAL_RACE
End Property

Public Property Let REAL_RACE(ByVal strX As String)
    strREAL_RACE = strX
End Property

Public Property Get PARTICIPANTS() As String
    PARTICIPANTS = strPARTICIPANTS
End Property

Public Property Let PARTICIPANTS(ByVal strX As String)
    strPARTICIPANTS = strX
End Property

Public Property Get RACE_NAME() As String
    RACE_NAME = strRACE_NAME
End Property

Public Property Let RACE_NAME(ByVal strX As String)
    strRACE_NAME = strX
End Property

Public Property Get RACE_YEAR() As String
    RACE_YEAR = strRACE_YEAR
End Property

Public Property Let RACE_YEAR(ByVal strX As String)
    strRACE_YEAR = strX
End Property

Public Property Get TRACK_LOCATION() As String
    TRACK_LOCATION = strTRACK_LOCATION
End Property

Public Property Let TRACK_LOCATION(ByVal strX As String)
    strTRACK_LOCATION = strX
End Property

Public Property Get COUNTRY() As String
    COUNTRY = strCOUNTRY
End Property

Public Property Let COUNTRY(ByVal strX As String)
    strCOUNTRY = strX
End Property

Public Property Get TRACK_NAME() As String
    TRACK_NAME = strTRACK_NAME
End Property

Public Property Let TRACK_NAME(ByVal strX As String)
    strTRACK_NAME = strX
End Property

Public Property Get TRACK_COLOUR() As Long
    TRACK_COLOUR = lngTRACK_COLOUR
End Property

Public Property Let TRACK_COLOUR(ByVal lngX As Long)
    lngTRACK_COLOUR = lngX
End Property

Public Property Get RACE_TYPE() As String
    RACE_TYPE = strRACE_TYPE
End Property

Public Property Let RACE_TYPE(ByVal strX As String)
    strRACE_TYPE = strX
    Select Case strX
        Case "F" 'Flat race
            strRACE_TYPE_TEXT = GetText(g_arr_Text, "RACETYPE01")
        Case "S" 'Steeplechase
            strRACE_TYPE_TEXT = GetText(g_arr_Text, "RACETYPE02")
        Case Else
    End Select
End Property

Public Property Get RACE_TYPE_TEXT() As String
    RACE_TYPE_TEXT = strRACE_TYPE_TEXT
End Property

Public Property Get METERS() As Integer
    METERS = intMETERS
End Property

Public Property Let METERS(ByVal intX As Integer)
    intMETERS = intX
End Property

Public Property Get STARTING_GATE() As String
    STARTING_GATE = strSTARTING_GATE
End Property

Public Property Let STARTING_GATE(ByVal strX As String)
    strSTARTING_GATE = strX
End Property

Public Property Get NUMBER_ENROLLED() As Integer
    NUMBER_ENROLLED = intNUMBER_ENROLLED
End Property

Public Property Let NUMBER_ENROLLED(ByVal intX As Integer)
    intNUMBER_ENROLLED = intX
End Property

Public Property Get LANES_FIX_OR_RANDOM() As String
    LANES_FIX_OR_RANDOM = strLANES_FIX_OR_RANDOM
End Property

Public Property Let LANES_FIX_OR_RANDOM(ByVal strX As String)
    strLANES_FIX_OR_RANDOM = strX
End Property

Public Property Get COLOURS_FIX_OR_RANDOM() As String
    COLOURS_FIX_OR_RANDOM = strCOLOURS_FIX_OR_RANDOM
End Property

Public Property Let COLOURS_FIX_OR_RANDOM(ByVal strX As String)
    strCOLOURS_FIX_OR_RANDOM = strX
End Property

Public Property Get ODDS_FIX_FIX_OR_RANDOM() As String
    ODDS_FIX_FIX_OR_RANDOM = strODDS_FIX_FIX_OR_RANDOM
End Property

Public Property Let ODDS_FIX_FIX_OR_RANDOM(ByVal strX As String)
    strODDS_FIX_FIX_OR_RANDOM = strX
End Property

Public Property Get ADVERTISING() As String
    ADVERTISING = strADVERTISING
End Property

Public Property Let ADVERTISING(ByVal strX As String)
    strADVERTISING = strX
End Property

Public Property Get BETTING_ALLOWED() As String
    BETTING_ALLOWED = strBETTING_ALLOWED
End Property

Public Property Let BETTING_ALLOWED(ByVal strX As String)
    strBETTING_ALLOWED = strX
End Property

Public Property Get TRACK_SURFACE() As String
    TRACK_SURFACE = strTRACK_SURFACE
End Property

Public Property Let TRACK_SURFACE(ByVal strX As String)
    strTRACK_SURFACE = strX
    Select Case strX
        Case "T" 'Turf track
            strTRACK_SURFACE_TEXT = GetText(g_arr_Text, "TRACK001")
            blnSQUIRT = False  'ToDo..... auslagern in GEMA Geläuf-Matrix
        Case "D" 'Dirt track
            strTRACK_SURFACE_TEXT = GetText(g_arr_Text, "TRACK002")
            blnSQUIRT = False
        Case "S" 'Snow track
            strTRACK_SURFACE_TEXT = GetText(g_arr_Text, "TRACK003")
            blnSQUIRT = False
        Case "A" 'Artificial track
            strTRACK_SURFACE_TEXT = GetText(g_arr_Text, "TRACK004")
            blnSQUIRT = False
        Case "M" 'Mudflats track
            strTRACK_SURFACE_TEXT = GetText(g_arr_Text, "TRACK005")
            blnSQUIRT = True
    End Select
End Property

Public Property Get TRACK_SURFACE_TEXT() As String
    TRACK_SURFACE_TEXT = strTRACK_SURFACE_TEXT
End Property

Public Property Get SQUIRT() As Boolean
    SQUIRT = blnSQUIRT
End Property

Public Property Get SELECTED() As String
    SELECTED = strRaceSelected
End Property

Public Property Let SELECTED(ByVal strX As String)
    strRaceSelected = strX
End Property

Public Property Get STARTED() As Boolean
    STARTED = blnRaceStarted
End Property

Public Property Let STARTED(ByVal blnX As Boolean)
    blnRaceStarted = blnX
End Property
