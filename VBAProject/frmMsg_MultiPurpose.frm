VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMsg_MultiPurpose 
   Caption         =   "[Caption]"
   ClientHeight    =   2568
   ClientLeft      =   72
   ClientTop       =   264
   ClientWidth     =   5136
   OleObjectBlob   =   "frmMsg_MultiPurpose.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmMsg_MultiPurpose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with a reusable message box
'   UserForm frmMsg_MultiPurpose

Private Sub UserForm_Initialize()
    'Display the pop-up in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

'Return the value of the button
'------------------------------
Private Sub cmdMsgOK_Click()
    g_enumButton = enumButton.OK
    Unload Me
End Sub

Private Sub cmdMsgCancel_Click()
    g_enumButton = enumButton.Cancel
    Unload Me
End Sub

Private Sub cmdMsgYes_Click()
    g_enumButton = enumButton.yes
    Unload Me
End Sub

Private Sub cmdMsgNo_Click()
    g_enumButton = enumButton.no
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Return "Cancel" if 'X' is clicked
    If CloseMode = 0 Then g_enumButton = enumButton.Cancel
End Sub
