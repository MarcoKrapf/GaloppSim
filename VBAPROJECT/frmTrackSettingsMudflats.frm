VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTrackSettingsMudflats 
   Caption         =   "UserForm1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6735
   OleObjectBlob   =   "frmTrackSettingsMudflats.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmTrackSettingsMudflats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    Call LabelCaptions
    
    With Me
    
        'Lugworm population (%)
        With .scrMud01
            .min = 0 'minumum value
            .max = 100 'maxmum value
            .SmallChange = 1 'value change when using the arrows
            .LargeChange = 25 'value change when clicking inside the slider
            .Value = objOption.LUGWORMS
        End With
        
        'Sea level (cm)
        With .scrMud02
            .min = 0 'minumum value
            .max = 10 'maxmum value
            .SmallChange = 1 'value change when using the arrows
            .LargeChange = 5 'value change when clicking inside the slider
            .Value = objOption.TIDE
        End With
        
    End With

    'Display the UserForm in the center of the Window
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
        .Caption = objRace.TRACK_SURFACE_TEXT
        .lblMud01.Caption = GetText(g_arr_Text, "TRACKSPEC001")
        .lblMud02.Caption = GetText(g_arr_Text, "TRACKSPEC002") & ": " & scrMud02.Value & GetText(g_arr_Text, "TRACKSPEC004")
        .lblMud03.Caption = GetText(g_arr_Text, "TRACKSPEC003") & ": " & scrMud01.Value & "%"
    End With
End Sub

'OK button
Private Sub cmdMud01_Click()
    objOption.LUGWORMS = scrMud01.Value
    objOption.TIDE = scrMud02.Value
    
    'Close UserForm
    Unload Me
End Sub

