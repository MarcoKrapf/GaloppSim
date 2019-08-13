VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOdds 
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10965
   OleObjectBlob   =   "frmOdds.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmOdds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the horses´ speed, condition and odds
'   UserForm frmOdds

Private Sub UserForm_Initialize()
    'Display the UserForm in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub
