VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRaceInfo 
   Caption         =   "[Race info]"
   ClientHeight    =   1620
   ClientLeft      =   -36
   ClientTop       =   -192
   ClientWidth     =   4476
   OleObjectBlob   =   "frmRaceInfo.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmRaceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with information during the race
'   UserForm frmRaceInfo

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = 1 'Prevent the pop-up from closing by clicking on "X" in the upper right corner
End Sub
