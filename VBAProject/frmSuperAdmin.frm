VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSuperAdmin 
   Caption         =   "Super Admin Tools"
   ClientHeight    =   2784
   ClientLeft      =   12
   ClientTop       =   72
   ClientWidth     =   5520
   OleObjectBlob   =   "frmSuperAdmin.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmSuperAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the Super Admin GUI
'   UserForm frmSuperAdmin

Private Sub UserForm_Initialize()
    With Me
        .width = 290
        .Height = 140
        .StartUpPosition = 2 'Display the UserForm in the center of the screen
    End With
End Sub

'Call the Test Suite GUI
Private Sub btnTestSuite_Click()
    frmTestSuite.show (vbModeless)
    Unload Me 'Close the Super Admin GUI
End Sub

'Call the Machine Learning Race Simulation GUI
Private Sub btnMLsimulation_Click()
    frmMachineLearning.show (vbModeless)
    Unload Me 'Close the Super Admin GUI
End Sub
