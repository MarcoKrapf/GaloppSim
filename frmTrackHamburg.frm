VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTrackHamburg 
   ClientHeight    =   3360
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10230
   OleObjectBlob   =   "frmTrackHamburg.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmTrackHamburg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    With Me
        .lblT01.Caption = txt(235)
        .lblT02.Caption = txt(236)
        .lblT03.Caption = txt(237) & " " & _
            txt(238) & " " & _
            txt(239) & " " & _
            txt(240) & " " & _
            txt(241)
    End With
End Sub
