VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRS_languages 
   ClientHeight    =   2628
   ClientLeft      =   12
   ClientTop       =   -72
   ClientWidth     =   5724
   OleObjectBlob   =   "frmRS_languages.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmRS_languages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the language selection
'   UserForm frmRS_languages

Private Sub UserForm_Initialize()
    With Me
        'Captions
        .caption = g_c_tool
        .Height = 110
        .width = 250
        .fraRS_lan01.caption = GetText(g_arr_Text, "LANGUAGE001")
        .optRS_lan01.caption = GetText(g_arr_Text, "LANGUAGE002") 'German
        .optRS_lan02.caption = GetText(g_arr_Text, "LANGUAGE003") 'English
        'OK button
        With cmdRS_lan01
            .caption = GetText(g_arr_Text, "BTN014")
        End With
        'Get the values of the radio buttons depending on the current language
        .optRS_lan01.Value = (objOption.language = "DE")
        .optRS_lan02.Value = (objOption.language = "EN")
    End With
    
    'Display the UserForm in the centre of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

'Click on the OK button
Private Sub cmdRS_lan01_Click()
    'Set the values of the radio buttons
    Select Case True
        Case optRS_lan01.Value
            objOption.language = "DE"
        Case optRS_lan02.Value
            objOption.language = "EN"
    End Select
    Unload Me 'Close the UserForm
End Sub
