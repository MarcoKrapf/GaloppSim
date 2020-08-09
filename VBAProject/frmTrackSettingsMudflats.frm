VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTrackSettingsMudflats 
   Caption         =   "[Race specific settings]"
   ClientHeight    =   3396
   ClientLeft      =   12
   ClientTop       =   -12
   ClientWidth     =   6288
   OleObjectBlob   =   "frmTrackSettingsMudflats.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmTrackSettingsMudflats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the track specific settings for a mudflats race
'   UserForm frmTrackSettingsMudflats

Private Sub UserForm_Initialize()

    With Me
        .Height = 185
        .width = 356
    End With
    
    Call LabelCaptions
    
    With Me
        'Lugworm population (%)
        With .scrMud01
            .min = 0 'Minumum value
            .max = 100 'Minumum value
            .SmallChange = 1 'Value change when using the arrows
            .LargeChange = 25 'Value change when clicking inside the slider
            .Value = objOption.LUGWORMS
        End With
        
        'Sea level (cm)
        With .scrMud02
            .min = 0 'Minumum value
            .max = 10 'Maximum value
            .SmallChange = 1 'Value change when using the arrows
            .LargeChange = 5 'Value change when clicking inside the slider
            .Value = objOption.TIDE
        End With
    End With

    'Display the UserForm in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
    
End Sub

'Lugworm population slider
Private Sub scrMud01_Change()
    Call LabelCaptions
End Sub

'Sea level slider
Private Sub scrMud02_Change()
    Call LabelCaptions
End Sub

'Label captions
Private Sub LabelCaptions()
    With Me
        .caption = GetText(g_arr_Text, "RACESPEC001")
        .lblMud01.caption = GetText(g_arr_Text, "TRACK005")
        .lblMud02.caption = GetText(g_arr_Text, "RACESPEC002") & ": " & scrMud02.Value & GetText(g_arr_Text, "RACESPEC004") 'Sea level (cm)
        .lblMud03.caption = GetText(g_arr_Text, "RACESPEC003") & ": " & scrMud01.Value & "%" 'Lugworm population (%)
    End With
End Sub

'Click on the OK button
Private Sub cmdMud01_Click()
    'Set the selected values
    objOption.LUGWORMS = scrMud01.Value
    objOption.TIDE = scrMud02.Value
    Unload Me 'Close the pop-up
End Sub
