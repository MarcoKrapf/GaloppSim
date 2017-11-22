VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStart 
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11115
   OleObjectBlob   =   "frmStart.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub boxBetSlips_Click()
    Dim i As Integer
    For i = 0 To boxBetSlips.ListCount - 1
        If boxBetSlips.Selected(i) Then
            Call Receipt(i + 1)
            Exit For
        End If
    Next i
End Sub

Private Sub btnS1_Click() 'Create new bet slip
    Call Gambler
End Sub

Private Sub btnS2_Click() 'Start race
    rennen = True
    Unload Me
End Sub

Private Sub btnS6_Click()
    Call odds
End Sub

Private Sub lblS4_Click()
    Select Case raceID
        Case "DD17"
            frmTrackHamburg.Show
        Case Else

    End Select
End Sub

Private Sub UserForm_Initialize()
    With Me
        .Caption = tool
    End With
    btnS2.SetFocus
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then 'UserForm closed
        If MsgBox(txt(103), vbCritical + vbYesNo, tool) = vbYes Then '>>yes<< clicked
            rennen = False
        Else '>>no<< clicked
            Cancel = 1
        End If
    Else 'start button pressed
        rennen = True
    End If
End Sub
