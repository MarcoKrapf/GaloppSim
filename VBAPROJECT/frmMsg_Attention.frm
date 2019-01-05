VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMsg_Attention 
   Caption         =   "[Caption]"
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16455
   OleObjectBlob   =   "frmMsg_Attention.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmMsg_Attention"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()

    Me.Caption = g_strMsgCaption
    
    'Adjust the text label
    With lblMsg01
        .top = 12
        .left = 12
        .Width = 800 'Initial width
        .Caption = g_strMsgText
        .TextAlign = fmTextAlignLeft
        .PicturePosition = fmPicturePositionLeftCenter
        .AutoSize = True
    End With
    
    'Adjust the size of the pop-up
    With Me
        .Width = lblMsg01.Width + 35
        .Height = lblMsg01.Height + 50
    End With
    
    'Display the UserForm in the center of the Window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

Private Sub UserForm_Click()
    Unload Me
End Sub

Private Sub lblMsg01_Click()
    Unload Me
End Sub
