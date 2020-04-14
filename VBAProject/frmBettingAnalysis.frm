VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBettingAnalysis 
   Caption         =   "[Caption]"
   ClientHeight    =   5448
   ClientLeft      =   -48
   ClientTop       =   -168
   ClientWidth     =   8988
   OleObjectBlob   =   "frmBettingAnalysis.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmBettingAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the betting analysis after the race
'   UserForm frmBettingAnalysis

Private Sub UserForm_Initialize()
    'Display the UserForm in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

