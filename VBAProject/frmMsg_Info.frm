VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMsg_Info 
   Caption         =   "[Caption]"
   ClientHeight    =   2268
   ClientLeft      =   84
   ClientTop       =   264
   ClientWidth     =   7224
   OleObjectBlob   =   "frmMsg_Info.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmMsg_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with a reusable info or warning message box
'   UserForm frmMsg_Info

Private Sub UserForm_Initialize()
    'Display the UserForm in the center of the Window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

'Close the pop-up on click on any item
'-------------------------------------
Private Sub UserForm_Click()
    Unload Me
End Sub

Private Sub imgAttention_Click()
    Unload Me
End Sub

Private Sub imgInformation_Click()
    Unload Me
End Sub

Private Sub lblText_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Unload Me
End Sub
