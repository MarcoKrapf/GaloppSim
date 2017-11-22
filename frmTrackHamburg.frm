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
        .lblT01.Caption = "Galopprennbahn Hamburg-Horn"
        .lblT02.Caption = "Geolocation: 53° 33' 31'' N, 10° 5' 7'' E"
        .lblT03.Caption = "Die Horner Rennbahn ist die größte in Deutschland. " & _
            "Auf der Grasbahn werden Flachrennen, aber auch Seejagdrennen gelaufen. " & _
            "Mehr als 50000 Zuschauer finden hier Platz, davon über 3000 auf der Tribüne. " & _
            "Das erste Rennen fand im Juli 1855 statt. " & _
            "Seit 1869 ist Hamburg-Horn der Austragungsort für das Deutsche Derby."
    End With
End Sub
