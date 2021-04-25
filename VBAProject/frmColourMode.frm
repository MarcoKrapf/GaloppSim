VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmColourMode 
   Caption         =   "[COLOUR MODE]"
   ClientHeight    =   6192
   ClientLeft      =   -348
   ClientTop       =   -1668
   ClientWidth     =   14904
   OleObjectBlob   =   "frmColourMode.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmColourMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Pop-up with the betting analysis after the race
'   UserForm frmColourMode

Private Sub UserForm_Initialize()
    'Display the UserForm in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
    
    'Captions
    With Me
        .caption = GetText(g_arr_Text, "COLOURS001")
        .lblStandard.caption = GetText(g_arr_Text, "COLOURS002")
        .lblTV1960.caption = GetText(g_arr_Text, "COLOURS003")
        .lblPopArt.caption = GetText(g_arr_Text, "COLOURS004")
        .lblPsychedelicArt.caption = GetText(g_arr_Text, "COLOURS005")
        .lblSmarties.caption = GetText(g_arr_Text, "COLOURS006")
        .lbl24h.caption = GetText(g_arr_Text, "COLOURS007")
        .lbl24hMorning.caption = GetText(g_arr_Text, "COLOURS009")
        .lbl24hEvening.caption = GetText(g_arr_Text, "COLOURS010")
        .lblDarkMode.caption = GetText(g_arr_Text, "COLOURS008")
        With .scr24h
            .min = -4
            .max = 4
            .SmallChange = 1
            .LargeChange = 1
            .Value = objOption.DAYLIGHT
        End With
        .Height = 315
        .width = 700
    End With
End Sub

Private Sub imgStandard_Click()
    Call ChangeColourMode("STANDARD")
    Unload Me 'Close the UserForm
End Sub

Private Sub lblStandard_Click()
    Call ChangeColourMode("STANDARD")
    Unload Me 'Close the UserForm
End Sub

Private Sub imgPopArt_Click()
    Call ChangeColourMode("POPART")
    Unload Me 'Close the UserFor
End Sub

Private Sub lblPopArt_Click()
    Call ChangeColourMode("POPART")
    Unload Me 'Close the UserFor
End Sub

Private Sub imgSmarties_Click()
    Call ChangeColourMode("SMARTIES")
    Unload Me 'Close the UserFor
End Sub

Private Sub lblSmarties_Click()
    Call ChangeColourMode("SMARTIES")
    Unload Me 'Close the UserFor
End Sub

Private Sub imgLSD_Click()
    Call ChangeColourMode("LSD")
    Unload Me 'Close the UserFor
End Sub

Private Sub lblPsychedelicArt_Click()
    Call ChangeColourMode("LSD")
    Unload Me 'Close the UserFor
End Sub

Private Sub imgTV1960_Click()
    Call ChangeColourMode("TV1960")
    Unload Me 'Close the UserFor
End Sub

Private Sub lblTV1960_Click()
    Call ChangeColourMode("TV1960")
    Unload Me 'Close the UserFor
End Sub

Private Sub imgDarkMode_Click()
    Call ChangeColourMode("DARKMODE")
    Unload Me 'Close the UserFor
End Sub

Private Sub lblDarkMode_Click()
    Call ChangeColourMode("DARKMODE")
    Unload Me 'Close the UserFor
End Sub

Private Sub img24h_Click()
    Call ChangeColourMode("24H")
    Unload Me 'Close the UserFor
End Sub

Private Sub lbl24h_Click()
    Call ChangeColourMode("24H")
    Unload Me 'Close the UserFor
End Sub
