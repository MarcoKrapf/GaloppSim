VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRS_navigation 
   Caption         =   "[Run Simple edition]"
   ClientHeight    =   1920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3960
   OleObjectBlob   =   "frmRS_navigation.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmRS_navigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNavigation_Click()
    g_wksRace.Activate 'activate the GALOPPSIM worksheet
    Call basAuxiliary.Scroll(1, 1) 'scroll to the upper left
    Unload Me
End Sub

Private Sub UserForm_Activate()
    With Me 'place the navigation panel in the upper left corner
        .top = Application.ActiveWindow.top + 20
        .left = Application.ActiveWindow.left + 20
    End With
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = GetTxt(g_arrTxt, "USERFORM005")
    With cmdNavigation
        .ForeColor = g_lngRaceInfoForeColour
        .BackColor = g_lngRaceInfoBackColour
        .Caption = GetTxt(g_arrTxt, "NAVI001")
        .WordWrap = True
    End With
End Sub
