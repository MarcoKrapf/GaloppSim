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
    'Prepare the development team page
        'Marco Matjes
        With img_info_team01
            .top = 6
            .Height = 144
            .left = 6
            .Width = 108
        End With
        With lbl_info_team01a
            .top = 156
            .Height = 18
            .left = 12
            .Width = 102
            .Font.Size = 12
        End With
        With lbl_info_team01b
            .top = 174
            .Height = 24
            .left = 12
            .Width = 102
        End With
        'Florian
        With img_info_team02
            .top = 6
            .Height = 144
            .left = 120
            .Width = 108
        End With
        With lbl_info_team02a
            .top = 156
            .Height = 18
            .left = 126
            .Width = 102
            .Font.Size = 12
        End With
        With lbl_info_team02b
            .top = 174
            .Height = 24
            .left = 126
            .Width = 102
        End With
        'Paul
        With img_info_team03
            .top = 6
            .Height = 144
            .left = 234
            .Width = 108
        End With
        With lbl_info_team03a
            .top = 156
            .Height = 18
            .left = 240
            .Width = 102
            .Font.Size = 12
        End With
        With lbl_info_team03b
            .top = 174
            .Height = 24
            .left = 240
            .Width = 102
        End With
        'Michael
        With img_info_team04
            .top = 6
            .Height = 144
            .left = 348
            .Width = 108
        End With
        With lbl_info_team04a
            .top = 156
            .Height = 18
            .left = 354
            .Width = 102
            .Font.Size = 12
        End With
        With lbl_info_team04b
            .top = 174
            .Height = 24
            .left = 354
            .Width = 102
        End With
        'Meike
        With img_info_team05
            .top = 210
            .Height = 144
            .left = 6
            .Width = 108
        End With
        With lbl_info_team05a
            .top = 360
            .Height = 18
            .left = 12
            .Width = 102
            .Font.Size = 12
        End With
        With lbl_info_team05b
            .top = 378
            .Height = 24
            .left = 12
            .Width = 102
        End With
        'Natalie
        With img_info_team06
            .top = 210
            .Height = 144
            .left = 120
            .Width = 108
        End With
        With lbl_info_team06a
            .top = 360
            .Height = 18
            .left = 126
            .Width = 102
            .Font.Size = 12
        End With
        With lbl_info_team06b
            .top = 378
            .Height = 24
            .left = 126
            .Width = 102
        End With
        'Atanas
        With img_info_team07
            .top = 210
            .Height = 144
            .left = 234
            .Width = 108
        End With
        With lbl_info_team07a
            .top = 360
            .Height = 18
            .left = 240
            .Width = 102
            .Font.Size = 12
        End With
        With lbl_info_team07b
            .top = 378
            .Height = 24
            .left = 240
            .Width = 102
        End With
        'Sam
        With img_info_team08
            .top = 210
            .Height = 144
            .left = 348
            .Width = 108
        End With
        With lbl_info_team08a
            .top = 360
            .Height = 18
            .left = 354
            .Width = 102
            .Font.Size = 12
        End With
        With lbl_info_team08b
            .top = 378
            .Height = 24
            .left = 354
            .Width = 102
        End With
    'Prepare the algorithms page
        With multiPage_algo
            .top = 18
            .Height = 162
            .left = 120
            .Visible = True
        End With
        chk_info_algorithms01.Value = objOption.STOP_ALG 'set the checkbox value
        With lbl_info_algorithms01
            .Caption = GetText(g_arr_Text, "ALGO04")
            .Font.Size = 20
            .Font.Bold = True
            .ForeColor = RGB(240, 240, 240)
            .TextAlign = fmTextAlignCenter
            .top = 30
            .Height = 140
            .left = 120
            .Width = 360
            .BackColor = 192
            .Visible = False
        End With
        'Expert pictures
        With img_info_algorithms01
            .top = 24
            .Height = 160
            .left = 6
            .Width = 108
            .Visible = True
        End With
        With img_info_algorithms02
            .top = 24
            .Height = 160
            .left = 6
            .Width = 108
            .Visible = False
        End With
        With img_info_algorithms03
            .top = 24
            .Height = 160
            .left = 6
            .Width = 108
            .Visible = False
        End With
        'Spy pictures
        With img_info_privacy01a
            .top = 90
            .Height = 162
            .left = 372
            .Width = 114
            .Visible = True
        End With
        With img_info_privacy01b
            .top = 90
            .Height = 162
            .left = 372
            .Width = 114
            .Visible = False
        End With
    'Display the UserForm in the center of the Window
        Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

'Page "Source code"
    'Click on Octocat
    Private Sub btn_info_code01_Click()
        Call OpenURL("https://github.com/MarcoKrapf/GaloppSim")
    End Sub

'Page "Contact"
    'Click on the e-mail button
    Private Sub btn_info_contact01_Click()
        Call SendMail
    End Sub
    
    'Click on the website address
    Private Sub lbl_info_contact01b_Click()
        Call OpenURL("https://galoppsim.racing/")
    End Sub
    
    'Click on the e-mail address
    Private Sub lbl_info_contact01c_Click()
        Call SendMail
    End Sub
    
    Private Sub SendMail()
        Dim objMail As Object 'Shell object for the e-mail
        On Error Resume Next
        Set objMail = CreateObject("Shell.Application")
        objMail.ShellExecute "mailto:" & g_c_email _
            & "&subject=" & g_c_tool
        On Error GoTo 0
    End Sub
    
    'Click on the "GALOPPSIM.RACING" button
    Private Sub btn_info_contact02_Click()
        Call OpenURL("https://galoppsim.racing/")
    End Sub

    'Click on the GaloppSim Facebook button
    Private Sub btn_info_contact03_Click()
        Call OpenURL("https://www.facebook.com/GaloppSim-2026661264317844/about/")
    End Sub

    'Click on the MIG Facebook button
    Private Sub btn_info_contact04_Click()
        Call OpenURL("https://www.facebook.com/Matjes-Institut-f%C3%BCr-Galoppsimulation-564551170578449/about/")
    End Sub

    'Click on the MIG Website button
    Private Sub btn_info_contact05_Click()
        Call OpenURL("http://matjes-institut.de/")
    End Sub

'Page "Donation"
    'Click on the Hero
    Private Sub btn_info_donation01_Click()
        Call OpenURL("https://www.grosse-hilfe.de/")
    End Sub

    'Click on the QR code
    Private Sub btn_info_donation02_Click()
        Call OpenURL("https://www.grosse-hilfe.de/spenden/spendenformular-jetzt-spenden.html")
    End Sub
    
    'Open the Website of the foundation in the standard browser
    Private Sub OpenURL(URL As String)
        On Error Resume Next
        ActiveWorkbook.FollowHyperlink Address:=URL
        On Error GoTo 0
    End Sub

'Checkbox "Stop algorithms"
Private Sub chk_info_algorithms01_Click()
    objOption.STOP_ALG = chk_info_algorithms01.Value
    If objOption.STOP_ALG = True Then
        multiPage_algo.Visible = False
        img_info_algorithms01.Visible = False
        img_info_algorithms02.Visible = False
        img_info_algorithms03.Visible = True
        lbl_info_algorithms01.Visible = True
    Else
        multiPage_algo.Visible = True
        Call MultiPage_algo_Change
    End If
End Sub

'MultiPage "Algorithms"
Private Sub MultiPage_algo_Change()
    If multiPage_algo.Value Mod 2 = 0 Then
        img_info_algorithms01.Visible = True
        img_info_algorithms02.Visible = False
        img_info_algorithms03.Visible = False
        lbl_info_algorithms01.Visible = False
    Else
        img_info_algorithms01.Visible = False
        img_info_algorithms02.Visible = True
        img_info_algorithms03.Visible = False
        lbl_info_algorithms01.Visible = False
    End If
End Sub

''Page "Privacy Policy"
'Private Sub img_info_privacy01a_click() 'click on the spy (dressed)
'    img_info_privacy01a.Visible = False
'    img_info_privacy01b.Visible = True
'End Sub
'
'Private Sub img_info_privacy01b_click() 'click on the spy
'    img_info_privacy01b.Visible = False
'    img_info_privacy01a.Visible = True
'End Sub
