VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRS_languages 
   ClientHeight    =   3165
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "frmRS_languages.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmRS_languages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    With Me
        'Captions
        .Caption = g_c_tool
        .fraRS_lan01.Caption = GetTxt(g_arrTxt, "LANGUAGE001")
        .optRS_lan01.Caption = GetTxt(g_arrTxt, "LANGUAGE002") 'German
        .optRS_lan02.Caption = GetTxt(g_arrTxt, "LANGUAGE003") 'English
        .optRS_lan03.Caption = GetTxt(g_arrTxt, "LANGUAGE006") 'Bulgarian
        'OK button
        With cmdRS_lan01
            .Caption = GetTxt(g_arrTxt, "BTN014") 'text
            .AutoSize = True 'size
            .left = Me.Width - .Width - 18 'position
        End With
        .optRS_lan01.Value = (g_strLanguage = "DE") 'True/False
        .optRS_lan02.Value = (g_strLanguage = "EN")
        .optRS_lan03.Value = (g_strLanguage = "BG")
    End With
    
    'Display the UserForm in the centre of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

Private Sub cmdRS_lan01_Click() 'OK button
    'Set values
    Select Case True
        Case optRS_lan01.Value
            g_strLanguage = "DE"
        Case optRS_lan02.Value
            g_strLanguage = "EN"
        Case optRS_lan03.Value
            g_strLanguage = "BG"
    End Select
    'Close UserForm
    Unload Me
End Sub
