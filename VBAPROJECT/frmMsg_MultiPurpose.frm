VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMsg_MultiPurpose 
   Caption         =   "[Caption]"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4905
   OleObjectBlob   =   "frmMsg_MultiPurpose.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmMsg_MultiPurpose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    Me.Caption = g_strMsgCaption
    
    'Set all buttons invisible
    cmdMsgOK.Visible = False
    cmdMsgCancel.Visible = False
    cmdMsgYes.Visible = False
    cmdMsgNo.Visible = False
    
    'Adjust the background label
    With lblMsg02
        .BackColor = &HFFFFFF 'white
        .Caption = ""
        .top = 0
        .left = 0
    End With
    
    'Adjust the text label
    With lblMsg01
        .BackColor = &HFFFFFF 'white
        .Caption = g_strMsgText
        .top = 12
        .left = 12
        .AutoSize = True
    End With
    
    'Adjust the size of the background label
    With lblMsg02
        .Height = lblMsg01.Height + 30
        .Width = lblMsg01.Width + 35
    End With

    'Adjust the buttons
        Select Case g_strMsgButtons
            Case "OK"
                Call AdjustButton(cmdMsgOK, GetTxt(g_arrTxt, "BTN014"), 0)
            Case "CancelOK"
                Call AdjustButton(cmdMsgOK, GetTxt(g_arrTxt, "BTN014"), 0)
                Call AdjustButton(cmdMsgCancel, GetTxt(g_arrTxt, "BTN015"), cmdMsgOK.Width + 5)
            Case "YesNo"
                Call AdjustButton(cmdMsgYes, GetTxt(g_arrTxt, "BTN016"), cmdMsgNo.Width + 5)
                Call AdjustButton(cmdMsgNo, GetTxt(g_arrTxt, "BTN017"), 0)
        End Select
        
    'Adjust the size of the pop-up
    With Me
        .Width = lblMsg01.Width + 35
        .Height = lblMsg01.Height + 105
    End With
    
    'Display the pop-up in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
    
End Sub

Private Sub AdjustButton(cmdButton As Object, strText As String, intCorrection As Integer)
    With cmdButton
        .Visible = True
        .Caption = strText
        .top = lblMsg01.top + lblMsg01.Height + 30
        .left = lblMsg01.left + lblMsg01.Width - .Width - intCorrection
    End With
End Sub

Private Sub cmdMsgOK_Click()
    g_strButtonPressed = "OK"
    Unload Me
End Sub

Private Sub cmdMsgCancel_Click()
    g_strButtonPressed = "CANCEL"
    Unload Me
End Sub

Private Sub cmdMsgYes_Click()
    g_strButtonPressed = "YES"
    Unload Me
End Sub

Private Sub cmdMsgNo_Click()
    g_strButtonPressed = "NO"
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
            Case "YesNo"
                g_strButtonPressed = "NO"
        End Select
    End If
End Sub
