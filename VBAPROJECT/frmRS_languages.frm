VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRS_languages 
   ClientHeight    =   2625
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
        .Caption = g_c_TOOL
        .fraRS_lan01.Caption = g_strTxt(229)
        .optRS_lan01.Caption = g_strTxt(230) 'German
        .optRS_lan02.Caption = g_strTxt(231) 'English
'        .optRS_lan03.Caption = g_strTxt(232) 'Swiss German
'        .optRS_lan04.Caption = g_strTxt(234) 'Bulgarian
        'OK button
        With cmdRS_lan01
            .Caption = g_strTxt(180) 'text
            .AutoSize = True 'size
            .left = Me.Width - .Width - 18 'position
        End With
        .optRS_lan01.Value = (g_strLanguage = "DE") 'True/False
        .optRS_lan02.Value = (g_strLanguage = "EN")
'        .optRS_lan03.Value = (g_strLanguage = "CH")
'        .optRS_lan04.Value = (g_strLanguage = "BG")
    End With
    
    'Display the UserForm in the centre of the window
    Call basMainCode.PlaceUserFormInCenter(Me)
End Sub

Private Sub cmdRS_lan01_Click() 'OK button
    'Set values
    Select Case True
        Case optRS_lan01.Value
            g_strLanguage = "DE"
        Case optRS_lan02.Value
            g_strLanguage = "EN"
'        Case optRS_lan03.Value
'            g_strLanguage = "CH"
'        Case optRS_lan04.Value
'            g_strLanguage = "BG"
    End Select
    'Close UserForm
    Unload Me
End Sub
