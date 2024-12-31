VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTrackSettingsDesert 
   Caption         =   "UserForm1"
   ClientHeight    =   3924
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6804
   OleObjectBlob   =   "frmTrackSettingsDesert.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmTrackSettingsDesert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the track specific settings for a desert race
'   UserForm frmTrackSettingsDesert

Private Sub UserForm_Initialize()

    With Me
        .Height = 155
        .width = 356
        .chkDesertDust = objOption.SAND_DUST
        .cmdDesert01.SetFocus
    End With
    
    Call LabelCaptions

    'Display the UserForm in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)

End Sub

'Label captions
Private Sub LabelCaptions()
    With Me
        .caption = GetText(g_arr_Text, "RACESPEC001")
        .lblDesert01.caption = GetText(g_arr_Text, "TRACK012")
        .chkDesertDust.caption = GetText(g_arr_Text, "RACESPEC036")
    End With
End Sub

'Click on the OK button
Private Sub cmdDesert01_Click()
    'Set the selected values
    objOption.SAND_DUST = chkDesertDust.Value
    Unload Me 'Close the pop-up
End Sub
