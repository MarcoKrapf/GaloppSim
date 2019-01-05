VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'VARIABLES
Private strLanguage As String 'Interface Language
Private blnStopAlgorithms As Boolean 'If true: Algorithms at work!
Private strExcelMode As String 'AI edition: Excel options - Excel mode, TV mode with menu strip or TV mode full-screen
Private byteTrackZoom As Byte 'Track zoom level (1=small 2=medium 3=large)
Private blnIncidentRefuse As Boolean 'If true: Horses can refuse to run
Private intTactics As Integer 'Racing tactics (0 = no tactics, 3 = three phases, 6 = six phases)
Private blnSlipstream As Boolean 'If true: Slipstream effect is active
Private blnSlipstreamDouble As Boolean 'If true: Double slipstream effect is selected
Private blnSlipstreamShow As Boolean 'If true: Slipstream effect is displayed graphically
Private blnFocusedRun As Boolean 'If true: Focused Run mode is active
Private intFocusedRun As Integer 'Number of the focused horse
Private blnHighlightFoc As Boolean 'If true: The horse in focus is highlighted during the race
Private blnHoofprints As Boolean 'If true: Hoof prints are displayed
Private blnHorseNamesLeft As Boolean 'If true: The names of the horses are permanently displayed at the left margin
Private blnHorseColoursLeft As Boolean 'If true: The colours of the horse are displayed on the left margin
Private blnHighlightFav As Boolean 'If true: The favorite is highlighted during the race
Private blnHorseNamesFinish As Boolean 'If true: The names of the horses are displayed in the destination
Private blnHorseNamesPhoto As Boolean 'If true: The names of the horses are displayed on the photo of the finish
Private blnPhotoBlackWhite As Boolean 'If true: The photo of the finish is displayed black-and-white
Private blnRankingColours As Boolean 'If true: The horse colours are displayed on the ranking list
Private blnRankingDelay As Boolean 'If true: Results are displayed on the ranking list bottom up with delay
Private blnBettingMode As Boolean 'If true: Bettings can be placed
Private blnBetsPlaced As Boolean 'If true: Bets have been placed
Private blnBettingAnalysis As Boolean 'If true: Betting analysis is performed automatically after the race
Private intTrackMetres As Integer 'Meteranzeige an der Rennbahn
Private blnRaceInformation As Boolean 'If true: Race information is displayed during the race
Private blnRaceInfoPopup As Boolean 'If true: Race information is displayed in a pop-up
Private blnRaceInfoWorksheet As Boolean 'If true: Race information is displayed on the worksheet
Private blnRaceInfoLeader As Boolean 'If true: Tthe name of the current leader is displayed
Private blnRaceInfoProgressBar As Boolean 'If true: A progress bar of the race distance is displayed
Private lngRaceInfoBackColour As Long 'Background colour of the race information
Private lngRaceInfoForeColour As Long 'Foreground colour of the race information
Private intGMPspeedfactor As Integer 'Multiplication factor for the race speed
Private blnSpeech As Boolean 'Speech output
Private intTide As Integer 'Wadden Sea: Tide height (cm)
Private intLugwormPopulation As Integer 'Wadden Sea: Lugworm density (%)
Private lngLugwormColour As Long 'Wadden Sea: Colour of the lugworms
Private lngPuddleColour As Long 'Wadden Sea: Colour of the puddles

'PROPERTIES
'----------

Public Property Get language() As String
    language = strLanguage
End Property

Public Property Let language(ByVal strX As String)
    strLanguage = strX
End Property

Public Property Get STOP_ALG() As Boolean
    STOP_ALG = blnStopAlgorithms
End Property

Public Property Let STOP_ALG(ByVal blnX As Boolean)
    blnStopAlgorithms = blnX
End Property

Public Property Get EXCEL_MODE() As String
    EXCEL_MODE = strExcelMode
End Property

Public Property Let EXCEL_MODE(ByVal strX As String)
    strExcelMode = strX
End Property

Public Property Get ZOOM_LEVEL() As Byte
    ZOOM_LEVEL = byteTrackZoom
End Property

Public Property Let ZOOM_LEVEL(ByVal byteX As Byte)
    byteTrackZoom = byteX
End Property

Public Property Get REFUSE_RUN() As Boolean
    REFUSE_RUN = blnIncidentRefuse
End Property

Public Property Let REFUSE_RUN(ByVal blnX As Boolean)
    blnIncidentRefuse = blnX
End Property

Public Property Get TACTICS() As Integer
    TACTICS = intTactics
End Property

Public Property Let TACTICS(ByVal intX As Integer)
    intTactics = intX
End Property

Public Property Get SLIPSTREAM() As Boolean
    SLIPSTREAM = blnSlipstream
End Property

Public Property Let SLIPSTREAM(ByVal blnX As Boolean)
    blnSlipstream = blnX
End Property

Public Property Get SLIPSTREAM_DBL() As Boolean
    SLIPSTREAM_DBL = blnSlipstreamDouble
End Property

Public Property Let SLIPSTREAM_DBL(ByVal blnX As Boolean)
    blnSlipstreamDouble = blnX
End Property

Public Property Get SLIPSTREAM_SHOW() As Boolean
    SLIPSTREAM_SHOW = blnSlipstreamShow
End Property

Public Property Let SLIPSTREAM_SHOW(ByVal blnX As Boolean)
    blnSlipstreamShow = blnX
End Property

Public Property Get FOCUSED_RUN() As Boolean
    FOCUSED_RUN = blnFocusedRun
End Property

Public Property Let FOCUSED_RUN(ByVal blnX As Boolean)
    blnFocusedRun = blnX
End Property

Public Property Get FOCUSED_NR() As Integer
    FOCUSED_NR = intFocusedRun
End Property

Public Property Let FOCUSED_NR(ByVal intX As Integer)
    intFocusedRun = intX
End Property
Public Property Get HIGHLIGHT_FOC() As Boolean
    HIGHLIGHT_FOC = blnHighlightFoc
End Property

Public Property Let HIGHLIGHT_FOC(ByVal blnX As Boolean)
    blnHighlightFoc = blnX
End Property

Public Property Get HOOFPRINTS() As Boolean
    HOOFPRINTS = blnHoofprints
End Property

Public Property Let HOOFPRINTS(ByVal blnX As Boolean)
    blnHoofprints = blnX
End Property

Public Property Get NAMES_LEFT() As Boolean
    NAMES_LEFT = blnHorseNamesLeft
End Property

Public Property Let NAMES_LEFT(ByVal blnX As Boolean)
    blnHorseNamesLeft = blnX
End Property

Public Property Get COLOURS_LEFT() As Boolean
    COLOURS_LEFT = blnHorseColoursLeft
End Property

Public Property Let COLOURS_LEFT(ByVal blnX As Boolean)
    blnHorseColoursLeft = blnX
End Property

Public Property Get HIGHLIGHT_FAV() As Boolean
    HIGHLIGHT_FAV = blnHighlightFav
End Property

Public Property Let HIGHLIGHT_FAV(ByVal blnX As Boolean)
    blnHighlightFav = blnX
End Property

Public Property Get NAMES_FINISH() As Boolean
    NAMES_FINISH = blnHorseNamesFinish
End Property

Public Property Let NAMES_FINISH(ByVal blnX As Boolean)
    blnHorseNamesFinish = blnX
End Property

Public Property Get NAMES_PHOTO() As Boolean
    NAMES_PHOTO = blnHorseNamesPhoto
End Property

Public Property Let NAMES_PHOTO(ByVal blnX As Boolean)
    blnHorseNamesPhoto = blnX
End Property

Public Property Get PHOTO_BW() As Boolean
    PHOTO_BW = blnPhotoBlackWhite
End Property

Public Property Let PHOTO_BW(ByVal blnX As Boolean)
    blnPhotoBlackWhite = blnX
End Property

Public Property Get RANKING_COL() As Boolean
    RANKING_COL = blnRankingColours
End Property

Public Property Let RANKING_COL(ByVal blnX As Boolean)
    blnRankingColours = blnX
End Property

Public Property Get RANKING_DELAY() As Boolean
    RANKING_DELAY = blnRankingDelay
End Property

Public Property Let RANKING_DELAY(ByVal blnX As Boolean)
    blnRankingDelay = blnX
End Property

Public Property Get BET_MODE() As Boolean
    BET_MODE = blnBettingMode
End Property

Public Property Let BET_MODE(ByVal blnX As Boolean)
    blnBettingMode = blnX
End Property

Public Property Get BET_PLACED() As Boolean
    BET_PLACED = blnBetsPlaced
End Property

Public Property Let BET_PLACED(ByVal blnX As Boolean)
    blnBetsPlaced = blnX
End Property

Public Property Get BET_ANALYSIS() As Boolean
    BET_ANALYSIS = blnBettingAnalysis
End Property

Public Property Let BET_ANALYSIS(ByVal blnX As Boolean)
    blnBettingAnalysis = blnX
End Property

Public Property Get METRES_DISPLAY() As Integer
    METRES_DISPLAY = intTrackMetres
End Property

Public Property Let METRES_DISPLAY(ByVal intX As Integer)
    intTrackMetres = intX
End Property

Public Property Get RACE_INFO() As Boolean
    RACE_INFO = blnRaceInformation
End Property

Public Property Let RACE_INFO(ByVal blnX As Boolean)
    blnRaceInformation = blnX
End Property

Public Property Get RACE_INFO_POP() As Boolean
    RACE_INFO_POP = blnRaceInfoPopup
End Property

Public Property Let RACE_INFO_POP(ByVal blnX As Boolean)
    blnRaceInfoPopup = blnX
End Property

Public Property Get RACE_INFO_WKS() As Boolean
    RACE_INFO_WKS = blnRaceInfoWorksheet
End Property

Public Property Let RACE_INFO_WKS(ByVal blnX As Boolean)
    blnRaceInfoWorksheet = blnX
End Property

Public Property Get RACE_INFO_LEADER() As Boolean
    RACE_INFO_LEADER = blnRaceInfoLeader
End Property

Public Property Let RACE_INFO_LEADER(ByVal blnX As Boolean)
    blnRaceInfoLeader = blnX
End Property

Public Property Get RACE_INFO_PROGRESS() As Boolean
    RACE_INFO_PROGRESS = blnRaceInfoProgressBar
End Property

Public Property Let RACE_INFO_PROGRESS(ByVal blnX As Boolean)
    blnRaceInfoProgressBar = blnX
End Property

Public Property Get RACE_INFO_COL_B() As Long
    RACE_INFO_COL_B = lngRaceInfoBackColour
End Property

Public Property Let RACE_INFO_COL_B(ByVal lngX As Long)
    lngRaceInfoBackColour = lngX
End Property

Public Property Get RACE_INFO_COL_F() As Long
    RACE_INFO_COL_F = lngRaceInfoForeColour
End Property

Public Property Let RACE_INFO_COL_F(ByVal lngX As Long)
    lngRaceInfoForeColour = lngX
End Property

Public Property Get SPEED_FACTOR() As Integer
    SPEED_FACTOR = intGMPspeedfactor
End Property

Public Property Let SPEED_FACTOR(ByVal intX As Integer)
    intGMPspeedfactor = intX
End Property

Public Property Get SPEECH() As Boolean
    SPEECH = blnSpeech
End Property

Public Property Let SPEECH(ByVal blnX As Boolean)
    blnSpeech = blnX
End Property

Public Property Get TIDE() As Integer
    TIDE = intTide
End Property

Public Property Let TIDE(ByVal intX As Integer)
    intTide = intX
End Property

Public Property Get LUGWORMS() As Integer
    LUGWORMS = intLugwormPopulation
End Property

Public Property Let LUGWORMS(ByVal intX As Integer)
    intLugwormPopulation = intX
End Property

Public Property Get LUGWORM_COL() As Long
    LUGWORM_COL = lngLugwormColour
End Property

Public Property Let LUGWORM_COL(ByVal lngX As Long)
    lngLugwormColour = lngX
End Property

Public Property Get PUDDLE_COL() As Long
    PUDDLE_COL = lngPuddleColour
End Property

Public Property Let PUDDLE_COL(ByVal lngX As Long)
    lngPuddleColour = lngX
End Property
