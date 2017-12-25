VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRSnavi 
   Caption         =   "Run Simple mode"
   ClientHeight    =   465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2265
   OleObjectBlob   =   "frmRSnavi.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmRSnavi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnNavigation_Click()
    With ActiveWindow
        .ScrollColumn = 1
        .ScrollRow = 1
    End With
    Unload Me
End Sub

Private Sub UserForm_Activate()
'    MsgBox "UserForm_Activate()"
    Call Position
End Sub

Private Sub Position()
    Me.top = ActiveWindow.height - 110
    Me.left = ActiveWindow.width - 165
'    If frmRSnavi.Visible = True Then MsgBox "I am in the lower right."
End Sub

Private Sub UserForm_Initialize()
    btnNavigation.Caption = txt(157)
'    MsgBox "UserForm_Initialize()"
'    Call Position
End Sub
