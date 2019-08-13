VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReceipt 
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4050
   OleObjectBlob   =   "frmReceipt.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmReceipt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with a betting slip receipt
'   UserForm frmReceipt

Private Sub UserForm_Initialize()
    'Display the UserForm in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

'Close the pop-up on click on any item
'-------------------------------------
Private Sub UserForm_Click()
    Unload Me
End Sub

Private Sub lblR1_Click()
    Unload Me
End Sub

Private Sub lblR2_Click()
    Unload Me
End Sub

Private Sub lblR3_Click()
    Unload Me
End Sub

Private Sub lblR4_Click()
    Unload Me
End Sub

Private Sub lblR5_Click()
    Unload Me
End Sub

Private Sub lblR6_Click()
    Unload Me
End Sub

Private Sub lblR7_Click()
    Unload Me
End Sub
