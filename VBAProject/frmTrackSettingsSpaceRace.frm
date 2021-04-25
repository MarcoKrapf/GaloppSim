VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTrackSettingsSpaceRace 
   Caption         =   "UserForm1"
   ClientHeight    =   7692
   ClientLeft      =   24
   ClientTop       =   -24
   ClientWidth     =   13788
   OleObjectBlob   =   "frmTrackSettingsSpaceRace.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmTrackSettingsSpaceRace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the settings for a space race
'   UserForm frmTrackSettingsSpaceRace

Dim arrPlanets() As Variant 'All planets
Dim strPlanetNr As String 'Number of the selected planet
Dim strPlanetName As String 'Selected planet
Dim strPlanetSurface As String 'Planet surface
Dim strPlanetSurfaceText As String 'Planet surface
Dim lngPlanetColour As Long 'Planet colour
Dim strAliens As String 'Selected alien behaviour

Private Sub UserForm_Initialize()

    arrPlanets = Array(imgSpace_Moon, imgSpace_Mars, imgSpace_Jupiter, imgSpace_Pluto, imgSpace_Saturn)
    Call LabelCaptions
    
    With Me
        .Height = 325
        .width = 500
        
        'Planet
        .optSpace01.Value = (objOption.SPACE_PLANET = enumPlanets.moon)
        .optSpace02.Value = (objOption.SPACE_PLANET = enumPlanets.mars)
        .optSpace03.Value = (objOption.SPACE_PLANET = enumPlanets.jupiter)
        .optSpace04.Value = (objOption.SPACE_PLANET = enumPlanets.pluto)
        .optSpace05.Value = (objOption.SPACE_PLANET = enumPlanets.saturn)

        Call PlanetPictureSize(arrPlanets) 'Place the planet pictures
        Call PlanetChange 'Show the picture and the data of the chosen planet
        
        'Race distance
        With .scrSpace02
            .min = 100 'Minumum value
            .max = 4000 'Minumum value
            .SmallChange = 100 'Value change when using the arrows
            .Value = objRace.METRES
        End With
        Call RaceDistance(.scrSpace02.Value)
        
        'Alien behaviour
        .optSpace06.Value = (objOption.SPACE_ALIENS = enumAliens.friendly)
        .optSpace07.Value = (objOption.SPACE_ALIENS = enumAliens.unfriendly)
        
        'Kidnapping rate slider
        With .scrSpace01
            .min = 1 'Minumum value
            .max = 3 'Minumum value
            .SmallChange = 1 'Value change when using the arrows
            .LargeChange = 1 'Value change when clicking inside the slider
            .Value = objOption.SPACE_KIDNAPPINGRATE
        End With
    End With

    'Display the UserForm in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
    
End Sub

'Change of the race distance slider
Private Sub scrSpace02_Change()
    Call RaceDistance(scrSpace02.Value)
End Sub

'Change of the kidnapping slider
Private Sub scrSpace01_Change()
    Call LabelCaptions
End Sub

'Label captions
Private Sub LabelCaptions()
    With Me
        .caption = ""
        .lblSpace01.caption = GetText(g_arr_Text, "RACESPEC005")
        .fraSpace01.caption = GetText(g_arr_Text, "RACESPEC006")
        .optSpace01.caption = GetText(g_arr_Text, "RACESPEC007")
        .optSpace02.caption = GetText(g_arr_Text, "RACESPEC008")
        .optSpace03.caption = GetText(g_arr_Text, "RACESPEC009")
        .optSpace04.caption = GetText(g_arr_Text, "RACESPEC010")
        .optSpace05.caption = GetText(g_arr_Text, "RACESPEC011")
        .lblSpace03.caption = GetText(g_arr_Text, "RACESPEC012") & ":"
        .lblSpace04.caption = GetText(g_arr_Text, "RACESPEC013") & ":"
        .lblSpace05.caption = GetText(g_arr_Text, "TRACK006") & ":"
        .fraSpace02.caption = GetText(g_arr_Text, "RACESPEC015")
        .optSpace06.caption = GetText(g_arr_Text, "RACESPEC016")
        .optSpace07.caption = GetText(g_arr_Text, "RACESPEC017")
        .lblSpace06.caption = GetText(g_arr_Text, "RACESPEC018")
        .lblSpace07.caption = GetText(g_arr_Text, "RACESPEC019")
        .lblSpace08.caption = GetText(g_arr_Text, "RACESPEC020")
        .cmdSpace01.caption = GetText(g_arr_Text, "BTN014")
        
    End With
End Sub

'Show all data of the planet
Private Sub PlanetChange()
    Select Case strPlanetNr
        Case enumPlanets.moon 'Moon
            Call PlanetVisibility(True, False, False, False, False)
            Call PlanetDescription(strPlanetName, "RACESPEC021", "RACESPEC022")
            Call RaceDistance(400)
            objSpeed.SPEED_LOOP_LOW = -1600
            objSpeed.SPEED_LOOP_HIGH = 2200
        Case enumPlanets.mars 'Mars
            Call PlanetVisibility(False, True, False, False, False)
            Call PlanetDescription(strPlanetName, "RACESPEC023", "RACESPEC024")
            Call RaceDistance(400)
            objSpeed.SPEED_LOOP_LOW = -1200
            objSpeed.SPEED_LOOP_HIGH = 1800
        Case enumPlanets.jupiter 'Jupiter
            Call PlanetVisibility(False, False, True, False, False)
            Call PlanetDescription(strPlanetName, "RACESPEC025", "RACESPEC026")
            Call RaceDistance(400)
            objSpeed.SPEED_LOOP_LOW = 0
            objSpeed.SPEED_LOOP_HIGH = 500
        Case enumPlanets.pluto 'Pluto
            Call PlanetVisibility(False, False, False, True, False)
            Call PlanetDescription(strPlanetName, "RACESPEC027", "RACESPEC028")
            Call RaceDistance(2000)
            objSpeed.SPEED_LOOP_LOW = -5000
            objSpeed.SPEED_LOOP_HIGH = 10000
        Case enumPlanets.saturn 'Saturn
            Call PlanetVisibility(False, False, False, False, True)
            Call PlanetDescription(strPlanetName, "RACESPEC029", "RACESPEC030")
            Call RaceDistance(400)
            objSpeed.SPEED_LOOP_LOW = -2800
            objSpeed.SPEED_LOOP_HIGH = 3200
    End Select
End Sub

'Display the selected planet
Private Sub PlanetVisibility(moon As Boolean, mars As Boolean, _
                jupiter As Boolean, pluto As Boolean, saturn As Boolean)
    imgSpace_Moon.Visible = moon
    imgSpace_Mars.Visible = mars
    imgSpace_Jupiter.Visible = jupiter
    imgSpace_Pluto.Visible = pluto
    imgSpace_Saturn.Visible = saturn
End Sub

'Display the planet specific data
Private Sub PlanetDescription(strPlanetName As String, strGravity As String, strWeather As String)
    fraSpace03.caption = strPlanetName
    lblSpace03a.caption = GetText(g_arr_Text, strGravity)
    lblSpace04a.caption = GetText(g_arr_Text, strWeather)
    lblSpace05a.caption = strPlanetSurfaceText
End Sub

'Set the size of al planet pictures
Private Sub PlanetPictureSize(planets As Variant)
    Dim planet As Variant
    For Each planet In planets
        With planet
        .top = 48
        .left = 126
        .Height = 96
        .width = 96
        End With
    Next
End Sub

'Race distance
Private Sub RaceDistance(intD)
    fraSpace04.caption = GetText(g_arr_Text, "RACE024") & ": " & intD & GetText(g_arr_Text, "RACE008")
    scrSpace02.Value = intD 'Position of the slider
End Sub

'Click on an option button for changing the planet
Private Sub optSpace01_Click() 'Moon
    strPlanetNr = enumPlanets.moon
    strPlanetName = GetText(g_arr_Text, "RACESPEC007")
    strPlanetSurface = "MOON"
    strPlanetSurfaceText = GetText(g_arr_Text, "TRACK007")
    lngPlanetColour = 10921638
    Call PlanetChange
End Sub
Private Sub optSpace02_Click() 'Mars
    strPlanetNr = enumPlanets.mars
    strPlanetName = GetText(g_arr_Text, "RACESPEC008")
    strPlanetSurface = "MARS"
    strPlanetSurfaceText = GetText(g_arr_Text, "TRACK008")
    lngPlanetColour = 5672164
    Call PlanetChange
End Sub

Private Sub optSpace03_Click() 'Jupiter
    strPlanetNr = enumPlanets.jupiter
    strPlanetName = GetText(g_arr_Text, "RACESPEC009")
    strPlanetSurface = "JUPITER"
    strPlanetSurfaceText = GetText(g_arr_Text, "TRACK009")
    lngPlanetColour = 13285804
    Call PlanetChange
End Sub

Private Sub optSpace04_Click() 'Pluto
    strPlanetNr = enumPlanets.pluto
    strPlanetName = GetText(g_arr_Text, "RACESPEC010")
    strPlanetSurface = "PLUTO"
    strPlanetSurfaceText = GetText(g_arr_Text, "TRACK010")
    lngPlanetColour = 7839921
    Call PlanetChange
End Sub

Private Sub optSpace05_Click() 'Saturn
    strPlanetNr = enumPlanets.saturn
    strPlanetName = GetText(g_arr_Text, "RACESPEC011")
    strPlanetSurface = "SATURN"
    strPlanetSurfaceText = GetText(g_arr_Text, "TRACK009")
    lngPlanetColour = 9094344
    Call PlanetChange
End Sub

'Click on an option button for changing the alien behaviour
Private Sub optSpace06_Click()
    strAliens = enumAliens.friendly
    scrSpace01.Visible = False
    lblSpace06.Visible = False
    lblSpace07.Visible = False
    lblSpace08.Visible = False
End Sub

Private Sub optSpace07_Click()
    strAliens = enumAliens.unfriendly
    scrSpace01.Visible = True
    lblSpace06.Visible = True
    lblSpace07.Visible = True
    lblSpace08.Visible = True
End Sub

'Click on the OK button
Private Sub cmdSpace01_Click()

    'Set the selected values
    objRace.COUNTRY_NAME = strPlanetName
    objOption.SPACE_PLANET = strPlanetNr
    objRace.TRACK_SURFACE = strPlanetSurface
    objRace.TRACK_COLOUR = lngPlanetColour
    objRace.METRES = scrSpace02.Value
    objOption.SPACE_ALIENS = strAliens
    objOption.SPACE_KIDNAPPINGRATE = scrSpace01.Value

    'Adapt the popup captions
    With frmStart
        .lblS2.caption = objRace.RACE_TYPE_TEXT & " " & GetText(g_arr_Text, "RACE007") & " " & objRace.METRES & " " & GetText(g_arr_Text, "RACE009") 'Race type and distance
        .lblS3.caption = objRace.TRACK_NAME & " " & GetText(g_arr_Text, "RACE002") & " " & objRace.TRACK_LOCATION & " (" & strPlanetName & ")" 'Race track, loaction and planet
        With .lblS4
            .caption = strPlanetSurfaceText
            .BackColor = objRace.TRACK_COLOUR
        End With
    End With
    
    Unload Me 'Close the pop-up
End Sub
