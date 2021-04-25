VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInfo 
   Caption         =   "[Caption]"
   ClientHeight    =   8184
   ClientLeft      =   -156
   ClientTop       =   -888
   ClientWidth     =   10824
   OleObjectBlob   =   "frmInfo.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with tool information
'   UserForm frmInfo

Private Sub UserForm_Initialize()
    'Prepare the development team page

'        'Comment in the following lines for getting an understanding of the object types
'        Debug.Print TypeName(img_info_team01)
'        Debug.Print TypeName(lbl_info_team01a)

        'Marco Matjes
        Call AlignObject(img_info_team01, 6, 144, 6, 108)
        Call AlignObject(lbl_info_team01a, 156, 18, 12, 102, 12)
        Call AlignObject(lbl_info_team01b, 174, 24, 12, 102)
'        'Alternative: Align each object without calling the general procedure for all types
'        With img_info_team01 'This is an "Image" control
'            .top = 6
'            .Height = 144
'            .left = 6
'            .Width = 108
'        End With
'        With lbl_info_team01a 'This is a "Label" control
'            .top = 156
'            .Height = 18
'            .left = 12
'            .Width = 102
'            .Font.Size = 12
'        End With
'        With lbl_info_team01b 'This is a "Label" control"
'            .top = 174
'            .Height = 24
'            .left = 12
'            .Width = 102
'        End With

        'Florian
        Call AlignObject(img_info_team02, 6, 144, 120, 108)
        Call AlignObject(lbl_info_team02a, 156, 18, 126, 102, 12)
        Call AlignObject(lbl_info_team02b, 174, 24, 126, 102)

        'Paul
        Call AlignObject(img_info_team03, 6, 144, 234, 108)
        Call AlignObject(lbl_info_team03a, 156, 18, 240, 102, 12)
        Call AlignObject(lbl_info_team03b, 174, 24, 240, 102)

        'Michael
        Call AlignObject(img_info_team07, 6, 144, 348, 108)
        Call AlignObject(lbl_info_team07a, 156, 18, 354, 102, 12)
        Call AlignObject(lbl_info_team07b, 174, 24, 354, 102)

        'Meike
        Call AlignObject(img_info_team05, 210, 144, 6, 108)
        Call AlignObject(lbl_info_team05a, 360, 18, 12, 102, 12)
        Call AlignObject(lbl_info_team05b, 378, 24, 12, 102)

        'Enno
        Call AlignObject(img_info_team09, 210, 144, 120, 108)
        Call AlignObject(lbl_info_team09a, 360, 18, 126, 102, 12)
        Call AlignObject(lbl_info_team09b, 378, 24, 126, 102)

        'Atanas
        Call AlignObject(img_info_team04, 210, 144, 234, 108)
        Call AlignObject(lbl_info_team04a, 360, 18, 240, 102, 12)
        Call AlignObject(lbl_info_team04b, 378, 24, 240, 102)
        
        'Farida
        Call AlignObject(img_info_team10, 210, 144, 348, 108)
        Call AlignObject(lbl_info_team10a, 360, 18, 354, 102, 12)
        Call AlignObject(lbl_info_team10b, 378, 24, 354, 102)

    'Prepare the algorithms page
        With multiPage_algo 'Align the MultiPage tabs for the algorithms
            .top = 18
            .Height = 162
            .left = 120
            .Visible = True
        End With
        With lbl_info_algorithms01 'Label with the algorithm deactivation text
            .caption = GetText(g_arr_Text, "ALGO04")
            .Font.size = 20
            .Font.Bold = True
            .ForeColor = RGB(240, 240, 240)
            .TextAlign = fmTextAlignCenter
            .top = 30
            .Height = 140
            .left = 120
            .width = 360
            .BackColor = 192
            .Visible = False
        End With
        chk_info_algorithms01.Value = objOption.STOP_ALG 'Set the algorithm checkbox status
        'Provide three expert pictures
        Call AlignObject(img_info_algorithms01, 24, 160, 6, 108, , False) 'Mouth open
        Call AlignObject(img_info_algorithms02, 24, 160, 6, 108, , True) 'Mouth wide open
        Call AlignObject(img_info_algorithms03, 24, 160, 6, 108, , True) 'Mouth closed

    'Prepare the privacy policy page
        Call AlignObject(img_info_privacy01a, 90, 162, 372, 114, , False, GetText(g_arr_Text, "PRIVACY03"))
        Call AlignObject(img_info_privacy01b, 90, 162, 372, 114, , True, GetText(g_arr_Text, "PRIVACY06"))

    'Display the UserForm in the center of the window
        Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

'Page "Code"
    'Click on Octocat
    Private Sub btn_info_code01_Click()
        Call OpenURL("https://github.com/MarcoKrapf/GaloppSim")
    End Sub

'Page "Donation"
    'Click on the Hero
    Private Sub btn_info_donation01_Click()
        Call OpenURL("https://www.grosse-hilfe.de/")
    End Sub

    'Click on the money bill
    Private Sub btn_info_donation02_Click()
        Call OpenURL("https://www.grosse-hilfe.de/spenden/spendenformular-jetzt-spenden.html")
    End Sub

'Page "Contact & Social media"
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
    'Determine the expert´s mouth shape depending on the algortithm that is selected
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

'Page "Privacy Policy"
'Since there is not Click for Images event provided in the Dropdown on the upper right
'just take the BeforeDragOver event and simply overwrite the procedure signature
Private Sub img_info_privacy01a_Click()
    If LCase(InputBox(GetText(g_arr_Text, "PRIVACY04"), "")) <> LCase(GetText(g_arr_Text, "PRIVACY05")) Then Exit Sub
    img_info_privacy01a.Visible = False
    img_info_privacy01b.Visible = True
End Sub

Private Sub img_info_privacy01b_Click()
'    img_info_privacy01a.Visible = True
'    img_info_privacy01b.Visible = False
    'TRUE = -1 // FALSE = 0 --> Alternative...
    img_info_privacy01b.Visible = img_info_privacy01b.Visible + 1
    img_info_privacy01a.Visible = img_info_privacy01a.Visible - 1
End Sub

'Place any UserForm control object (like image or label) using the general "control" type for the different objects types
Private Sub AlignObject(obj As control, t As Integer, h As Integer, l As Integer, w As Integer, Optional fs As Integer, Optional hide As Boolean, Optional tip As String)
    With obj
        .top = t 'Top of the object
        .Height = h 'Height of the object
        .left = l 'Left of the object
        .width = w 'Width of the object
        If fs > 0 Then .Font.size = fs 'Only if the optional argument for the font size has been supplied
        If hide = True Then
            .Visible = False
        Else
            .Visible = True
        End If
        If tip <> "" Then .ControlTipText = tip 'Only if the optional argument for the control tip text has been supplied
    End With
End Sub

''Example: This would only work for images due to the use of the "Image" control type
'Private Sub AlignImage(obj As Image, t As Integer, h As Integer, l As Integer, w As Integer)
'    With obj
'        .top = t
'        .Height = h
'        .left = l
'        .Width = w
'    End With
'End Sub
