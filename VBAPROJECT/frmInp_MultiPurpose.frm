VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInp_MultiPurpose 
   Caption         =   "[Caption]"
   ClientHeight    =   1215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "frmInp_MultiPurpose.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmInp_MultiPurpose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    
    'Set all buttons invisible
    cmdInpOK.Visible = False
    cmdInpCancel.Visible = False
    
    'Adjust the text label
    With lblInp01
        .Caption = g_strMsgText
        .top = 12
        .left = 12
        .AutoSize = True
    End With
    
    'Adjust the text box
    With txtInp01
        .left = lblInp01.Width + 12 + 6
        .Width = 120
        .Height = 20
        .MaxLength = 26 'up to 26 letters allowed
    End With

    'Adjust the buttons
        Select Case g_strMsgButtons
            Case "OK"
                Call AdjustButton(cmdInpOK, GetText(g_arr_Text, "BTN014"), 0)
            Case "CancelOK"
                Call AdjustButton(cmdInpOK, GetText(g_arr_Text, "BTN014"), 0)
                Call AdjustButton(cmdInpCancel, GetText(g_arr_Text, "BTN015"), cmdInpOK.Width + 5)
        End Select

    'Adjust the size of the pop-up
    With Me
        .Width = txtInp01.left + txtInp01.Width + 20
        .Height = lblInp01.Height + 100
    End With
    
    'Display the pop-up in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

Private Sub AdjustButton(cmdButton As Object, strText As String, intCorrection As Integer)
    With cmdButton
        .Visible = True
        .Caption = strText
        .top = txtInp01.top + txtInp01.Height + 18
        .left = txtInp01.left + txtInp01.Width - .Width - intCorrection
    End With
End Sub

Private Sub cmdInpOK_Click()
    g_strButtonPressed = "OK"
    g_strPlayerName = txtInp01.Value
    Unload Me
End Sub

Private Sub cmdInpCancel_Click()
    g_strButtonPressed = "CANCEL"
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Determine standard button if 'X' is clicked
    If CloseMode = 0 Then
        Select Case g_strMsgButtons
            Case "OK"
                g_strButtonPressed = "OK"
            Case "CancelOK"
                g_strButtonPressed = "CANCEL"
        End Select
    End If
End Sub

