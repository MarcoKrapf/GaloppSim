VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInp_MultiPurpose 
   Caption         =   "[Caption]"
   ClientHeight    =   2640
   ClientLeft      =   24
   ClientTop       =   84
   ClientWidth     =   4992
   OleObjectBlob   =   "frmInp_MultiPurpose.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmInp_MultiPurpose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with a reusable input box
'   UserForm frmInp_MultiPurpose

Private Sub UserForm_Initialize()
    'Display the pop-up in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

'Click on the "OK" button
Private Sub cmdInpOK_Click()
    g_enumButton = enumButton.OK 'Value of the button
    g_strInpBoxReturnValue = txtInp01.Value 'Value of the input field
    Unload Me
End Sub

'Click on the "Cancel" button
Private Sub cmdInpCancel_Click()
    g_enumButton = enumButton.Cancel 'Value of the button
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Return "Cancel" if 'X' is clicked
    If CloseMode = 0 Then g_enumButton = enumButton.Cancel 'Value of the button
End Sub
