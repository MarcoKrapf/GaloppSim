VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTrackSettingsAirQuality 
   Caption         =   "UserForm1"
   ClientHeight    =   2856
   ClientLeft      =   -12
   ClientTop       =   -84
   ClientWidth     =   5856
   OleObjectBlob   =   "frmTrackSettingsAirQuality.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmTrackSettingsAirQuality"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the track specific settings for a race with
'particulates in the air
'   UserForm frmTrackSettingsAirQuality

Private Sub UserForm_Initialize()

    With Me
        .Height = 186
        .width = 350
    End With

    Call LabelCaptions
    
    'Particulates slider (PM10)
    With Me.scrDust
        .min = 0 'Minumum value
        .max = 5 'Minumum value
        .SmallChange = 1 'Value change when using the arrows
        .LargeChange = 1 'Value change when clicking inside the slider
        .Value = objOption.PARTICULATES_SLIDER
    End With

    'Display the UserForm in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
    
End Sub

'Particulates slider
Private Sub scrDust_Change()
    Call LabelCaptions
End Sub

'Label captions
Private Sub LabelCaptions()
    With Me
        .caption = GetText(g_arr_Text, "RACESPEC001")
        .lblDust01.caption = GetText(g_arr_Text, "RACESPEC031")
        .lblDust02.caption = GetText(g_arr_Text, "RACESPEC032")
        .lblDust03.caption = 10 + 40 * .scrDust.Value _
                & " " & GetText(g_arr_Text, "RACESPEC033") 'Particulates (my/m3)
    End With
End Sub

'Click on the OK button
Private Sub cmdDust_Click()
    'Set the selected values
    objOption.PARTICULATES_SLIDER = scrDust.Value
    Unload Me 'Close the pop-up
End Sub
