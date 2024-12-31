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

'Class Module for creating exactly one object instance for storing the race options
'   Class Module clsOptions

'OBJECT VARIABLES
Private strLanguage As String 'GUI language
Private blnStopAlgorithms As Boolean 'If true: Algorithms at work!
Private strExcelMode As String 'AI edition: Excel options - Excel mode, TV mode with menu strip or TV mode full-screen
Private blnIncidentRefuse As Boolean 'If true: Horses can refuse to run
Private intIncidentRefusalRate As Integer 'Start refusal rate
Private blnTactics As Boolean 'Racing tactics (true or false)
Private blnTacticsRevelationTac As Boolean 'If true: Reveal tactics
Private blnTacticsRevelationCurr As Boolean 'If true: Show the current speed
Private intSlipstream As Integer 'Impact of slipstream effect
Private blnSlipstreamShow As Boolean 'If true: Slipstream effect is displayed graphically
Private lngFocusedRun As Long 'Camera mode (enumCamera "standard", "focus_horse", "focus_leader")
Private intFocusedRun As Integer 'Number of the focused horse
Private blnHighlightFoc As Boolean 'If true: The horse in focus is highlighted during the race
Private blnHoofprints As Boolean 'If true: Hoof prints are displayed
Private blnHorseNamesLeft As Boolean 'If true: The names of the horses are permanently displayed at the left margin
Private blnHorseColoursLeft As Boolean 'If true: The colours of the horse are permanently displayed on the left margin
Private blnAnnounceFav As Boolean 'If true: The horses with the best chances will be announced before the race
Private blnHighlightFav As Boolean 'If true: The favourite horse is highlighted during the race
Private blnHorseNamesFinish As Boolean 'If true: The names of the horses are displayed in the finish area
Private blnHorseNamesPhoto As Boolean 'If true: The names of the horses are displayed on the photo of the finish
Private blnPhotoBlackWhite As Boolean 'If true: The photo of the finish is displayed black-and-white
Private blnRankingColours As Boolean 'If true: The horse colours are displayed on the ranking list
Private blnRankingDelay As Boolean 'If true: Results are displayed on the ranking list bottom up with delay
Private blnBettingMode As Boolean 'If true: Bettings can be placed
Private blnBetsPlaced As Boolean 'If true: Bets have been placed
Private blnBettingAnalysis As Boolean 'If true: Betting analysis is performed automatically after the race
Private intTrackMetres As Integer 'Track metres markers on the race track
Private blnRaceInformation As Boolean 'If true: Race information is displayed during the race
Private blnRaceInfoPopup As Boolean 'If true: Race information is displayed in a pop-up
Private blnRaceInfoWorksheet As Boolean 'If true: Race information is displayed on the Worksheet
Private blnRaceInfoLeader As Boolean 'If true: Tthe name of the current leader is displayed
Private blnRaceInfoProgressBar As Boolean 'If true: A progress bar of the race distance is displayed
Private lngRaceInfoBackColour As Long 'Background colour of the race information
Private lngRaceInfoForeColour As Long 'Font colour of the race information
Private intGMPspeedfactor As Integer 'Multiplication factor for the race speed
Private blnMomentumBars As Boolean 'If true: Show the Momentum speed bars
Private blnMomentumIcons As Boolean 'If true: Show the Momentum speed icons
Private intMomentumRefreshRate 'Refresh rate of the speed bars
Private blnRaceSpeedMonitor As Boolean 'If true: Show the race speed chart
Private blnRaceSpeedMonitorSpeed As Boolean 'If true: Show the current speed
Private blnRaceSpeedMonitorDistance As Boolean 'If true: Show the metres run
Private intRaceSpeedMonitorRefreshRate 'Refresh rate of the race speed monitor
Private blnSpeech As Boolean 'Speech output
Private blnSplash As Boolean 'Wadden Sea: Water splashes
Private intTide As Integer 'Wadden Sea: Tide height (cm)
Private intLugwormPopulation As Integer 'Wadden Sea: Lugworm density (%)
Private lngLugwormColour As Long 'Wadden Sea: Colour of the lugworms
Private lngPuddleColour As Long 'Wadden Sea: Colour of the puddles
Private strSpacePlanet As String 'Space race: Planet
Private strSpaceAliens As String 'Space race: Alien behaviour
Private intSpaceKidnapping As Integer 'Space race: Kidnapping rate
Private blnTribunes As Boolean 'If true: Draw tribunes, food and beverage
Private intSpectators As Integer 'Spectators (0 = none, 100 = sold out)
Private intParticulates As Integer 'Particulates (PM10)
Private lngParticulatesPattern As Long 'Particulates (PM10)
Private lngRaceSheetBackground As Long 'Background colour (behind the track)
Private lngRaceDataFontColour As Long 'Font colour of the general text
Private lngRaceScoreboardColour As Long 'Background colour of the ranking list
Private int24hDaylight As Integer 'Daylight factor used for 24h colour mode
Private lng24hDaylightCol As Long 'Daylight colour used for 24h colour mode
Private blnAutoFit As Boolean 'Auto-fit of the content according to the windows size
Private blnAutoSave As Boolean 'Save the race automatically to a file
Private blnStartingGridIn As Boolean 'Horses are already in the starting grid
Private blnStartingGridBehind As Boolean 'Horses are waiting behind the starting grid
Private blnSandDust As Boolean 'Desert: Sand dust clouds behind the camels
Private intTunnelCount As Integer 'Tunnels: Number of tunnels per row
Private intTunnelLength As Integer 'Tunnels: Total tunnel length per row
Private blnRoadCrossing As Boolean 'Road crossing the race track
Private blnPoliceCar As Boolean 'Race is secured by the police


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

Public Property Get REFUSE_RUN() As Boolean
    REFUSE_RUN = blnIncidentRefuse
End Property

Public Property Let REFUSE_RUN(ByVal blnX As Boolean)
    blnIncidentRefuse = blnX
End Property

Public Property Get REFUSAL_RATE() As Integer
    REFUSAL_RATE = intIncidentRefusalRate
End Property

Public Property Let REFUSAL_RATE(ByVal intX As Integer)
    intIncidentRefusalRate = intX
End Property

Public Property Get TACTICS() As Boolean
    TACTICS = blnTactics
End Property

Public Property Let TACTICS(ByVal blnX As Boolean)
    blnTactics = blnX
End Property

Public Property Get TACTICS_REVEAL_TAC() As Boolean
    TACTICS_REVEAL_TAC = blnTacticsRevelationTac
End Property

Public Property Let TACTICS_REVEAL_TAC(ByVal blnX As Boolean)
    blnTacticsRevelationTac = blnX
End Property

Public Property Get TACTICS_REVEAL_CURR() As Boolean
    TACTICS_REVEAL_CURR = blnTacticsRevelationCurr
End Property

Public Property Let TACTICS_REVEAL_CURR(ByVal blnX As Boolean)
    blnTacticsRevelationCurr = blnX
End Property

Public Property Get SLIPSTREAM_IMPACT() As Integer
    SLIPSTREAM_IMPACT = intSlipstream
End Property

Public Property Let SLIPSTREAM_IMPACT(ByVal intX As Integer)
    intSlipstream = intX
End Property

Public Property Get SLIPSTREAM_SHOW() As Boolean
    SLIPSTREAM_SHOW = blnSlipstreamShow
End Property

Public Property Let SLIPSTREAM_SHOW(ByVal blnX As Boolean)
    blnSlipstreamShow = blnX
End Property

Public Property Get FOCUSED_RUN() As Long
    FOCUSED_RUN = lngFocusedRun
End Property

Public Property Let FOCUSED_RUN(ByVal lngX As Long)
    lngFocusedRun = lngX
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

Public Property Get ANNOUNCE_FAV() As Boolean
    ANNOUNCE_FAV = blnAnnounceFav
End Property

Public Property Let ANNOUNCE_FAV(ByVal blnX As Boolean)
    blnAnnounceFav = blnX
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

Public Property Get MOMENTUM_BARS() As Boolean
    MOMENTUM_BARS = blnMomentumBars
End Property

Public Property Let MOMENTUM_BARS(ByVal blnX As Boolean)
    blnMomentumBars = blnX
End Property

Public Property Get MOMENTUM_ICONS() As Boolean
    MOMENTUM_ICONS = blnMomentumIcons
End Property

Public Property Let MOMENTUM_ICONS(ByVal blnX As Boolean)
    blnMomentumIcons = blnX
End Property

Public Property Get MOMENTUM_REFRESHRATE() As Integer
    MOMENTUM_REFRESHRATE = intMomentumRefreshRate
End Property

Public Property Let MOMENTUM_REFRESHRATE(ByVal intX As Integer)
    intMomentumRefreshRate = intX
End Property

Public Property Get SPEEDMONITOR() As Boolean
    SPEEDMONITOR = blnRaceSpeedMonitor
End Property

Public Property Let SPEEDMONITOR(ByVal blnX As Boolean)
    blnRaceSpeedMonitor = blnX
End Property

Public Property Get RSMON_SPEED() As Boolean
    RSMON_SPEED = blnRaceSpeedMonitorSpeed
End Property

Public Property Let RSMON_SPEED(ByVal blnX As Boolean)
    blnRaceSpeedMonitorSpeed = blnX
End Property

Public Property Get RSMON_DISTANCE() As Boolean
    RSMON_DISTANCE = blnRaceSpeedMonitorDistance
End Property

Public Property Let RSMON_DISTANCE(ByVal blnX As Boolean)
    blnRaceSpeedMonitorDistance = blnX
End Property

Public Property Get SPEEDMON_REFRESHRATE() As Integer
    SPEEDMON_REFRESHRATE = intRaceSpeedMonitorRefreshRate
End Property

Public Property Let SPEEDMON_REFRESHRATE(ByVal intX As Integer)
    intRaceSpeedMonitorRefreshRate = intX
End Property

Public Property Get SPEECH() As Boolean
    SPEECH = blnSpeech
End Property

Public Property Let SPEECH(ByVal blnX As Boolean)
    blnSpeech = blnX
End Property

Public Property Get WATER_SPLASHES() As Boolean
    WATER_SPLASHES = blnSplash
End Property

Public Property Let WATER_SPLASHES(ByVal blnX As Boolean)
    blnSplash = blnX
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

Public Property Get SPACE_PLANET() As String
    SPACE_PLANET = strSpacePlanet
End Property

Public Property Let SPACE_PLANET(ByVal strX As String)
    strSpacePlanet = strX
End Property

Public Property Get SPACE_ALIENS() As String
    SPACE_ALIENS = strSpaceAliens
End Property

Public Property Let SPACE_ALIENS(ByVal strX As String)
    strSpaceAliens = strX
End Property

Public Property Get SPACE_KIDNAPPINGRATE() As Integer
    SPACE_KIDNAPPINGRATE = intSpaceKidnapping
End Property

Public Property Let SPACE_KIDNAPPINGRATE(ByVal intX As Integer)
    intSpaceKidnapping = intX
End Property

Public Property Get TRIBUNES() As Boolean
    TRIBUNES = blnTribunes
End Property

Public Property Let TRIBUNES(ByVal blnX As Boolean)
    blnTribunes = blnX
End Property

Public Property Get SPECTATORS() As Integer
    SPECTATORS = intSpectators
End Property

Public Property Let SPECTATORS(ByVal intX As Integer)
    intSpectators = intX
End Property

Public Property Get PARTICULATES_SLIDER() As Integer
    PARTICULATES_SLIDER = intParticulates
End Property

Public Property Let PARTICULATES_SLIDER(ByVal intX As Integer)
    intParticulates = intX
End Property

Public Property Get PARTICULATES_PATTERN() As Integer
    PARTICULATES_PATTERN = lngParticulatesPattern
End Property

Public Property Let PARTICULATES_PATTERN(ByVal intX As Integer)
    lngParticulatesPattern = intX
End Property

Public Property Get COL_BACK() As Long
    COL_BACK = lngRaceSheetBackground
End Property

Public Property Let COL_BACK(ByVal lngX As Long)
    lngRaceSheetBackground = lngX
End Property

Public Property Get COL_TEXT() As Long
    COL_TEXT = lngRaceDataFontColour
End Property

Public Property Let COL_TEXT(ByVal lngX As Long)
    lngRaceDataFontColour = lngX
End Property

Public Property Get COL_RANKINGS() As Long
    COL_RANKINGS = lngRaceScoreboardColour
End Property

Public Property Let COL_RANKINGS(ByVal lngX As Long)
    lngRaceScoreboardColour = lngX
End Property

Public Property Get DAYLIGHT() As Integer
    DAYLIGHT = int24hDaylight
End Property

Public Property Let DAYLIGHT(ByVal intX As Integer)
    int24hDaylight = intX
    Select Case intX
        Case -4
            lng24hDaylightCol = 3546898
        Case -3
            lng24hDaylightCol = 5911068
        Case -2
            lng24hDaylightCol = 10508594
        Case -1
            lng24hDaylightCol = 15254622
        Case 0
            lng24hDaylightCol = 15784084
        Case 1
            lng24hDaylightCol = 15519612
        Case 2
            lng24hDaylightCol = 12940359
        Case 3
            lng24hDaylightCol = 5181281
        Case 4
            lng24hDaylightCol = 3546898
    End Select
End Property

Public Property Get DAYLIGHT_COL() As Long
    DAYLIGHT_COL = lng24hDaylightCol
End Property

Public Property Get AUTOFIT() As Boolean
    AUTOFIT = blnAutoFit
End Property

Public Property Let AUTOFIT(ByVal blnX As Boolean)
    blnAutoFit = blnX
End Property

Public Property Get AUTO_SAVE() As Boolean
    AUTO_SAVE = blnAutoSave
End Property

Public Property Let AUTO_SAVE(ByVal blnX As Boolean)
    blnAutoSave = blnX
End Property

Public Property Get STARTING_GRID_IN() As Boolean
    STARTING_GRID_IN = blnStartingGridIn
End Property

Public Property Let STARTING_GRID_IN(ByVal blnX As Boolean)
    blnStartingGridIn = blnX
End Property

Public Property Get STARTING_GRID_BEHIND() As Boolean
    STARTING_GRID_BEHIND = blnStartingGridBehind
End Property

Public Property Let STARTING_GRID_BEHIND(ByVal blnX As Boolean)
    blnStartingGridBehind = blnX
End Property

Public Property Get SAND_DUST() As Boolean
    SAND_DUST = blnSandDust
End Property

Public Property Let SAND_DUST(ByVal blnX As Boolean)
    blnSandDust = blnX
End Property

Public Property Get TUNNEL_COUNT() As Integer
    TUNNEL_COUNT = intTunnelCount
End Property

Public Property Let TUNNEL_COUNT(ByVal intX As Integer)
    intTunnelCount = intX
End Property

Public Property Get TUNNEL_LENGTH() As Integer
    TUNNEL_LENGTH = intTunnelLength
End Property

Public Property Let TUNNEL_LENGTH(ByVal intX As Integer)
    intTunnelLength = intX
End Property

Public Property Get ROADCROSSING() As Boolean
    ROADCROSSING = blnRoadCrossing
End Property

Public Property Let ROADCROSSING(ByVal blnX As Boolean)
    blnRoadCrossing = blnX
End Property

Public Property Get POLICECAR() As Boolean
    POLICECAR = blnPoliceCar
End Property

Public Property Let POLICECAR(ByVal blnX As Boolean)
    blnPoliceCar = blnX
End Property
