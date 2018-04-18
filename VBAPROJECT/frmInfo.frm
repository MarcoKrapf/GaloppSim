VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInfo 
   Caption         =   "[Caption]"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10245
   OleObjectBlob   =   "frmInfo.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    'Display the UserForm in the center of the Window
    Call basMainCode.PlaceUserFormInCenter(Me)
End Sub

'Page "Source code"
    'Click on Octocat
    Private Sub btn_info_sourcecode01_Click()
        On Error Resume Next
        ActiveWorkbook.FollowHyperlink Address:="https://github.com/MarcoKrapf/GaloppSim"
        On Error GoTo 0
    End Sub

'Page "Contact"
    'Click on the mail button
    Private Sub btn_info_contact01_Click()
        Dim objMail As Object 'Shell object for the e-mail
        On Error Resume Next
        Set objMail = CreateObject("Shell.Application")
        objMail.ShellExecute "mailto:" & g_c_email _
            & "&subject=" & g_c_TOOL
        On Error GoTo 0
    End Sub

'Page "Donation"
    'Click on the Hero
    Private Sub btn_info_donation01_Click()
        Call OpenGHFKH("https://www.grosse-hilfe.de/")
    End Sub

    'Click on the QR code
    Private Sub btn_info_donation02_Click()
        Call OpenGHFKH("http://spenden.grosse-hilfe.de/")
    End Sub
    
    'Open the Website of the foundation in the standard browser
    Private Sub OpenGHFKH(URL As String)
        On Error Resume Next
        ActiveWorkbook.FollowHyperlink Address:=URL
        On Error GoTo 0
    End Sub
