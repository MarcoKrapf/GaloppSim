VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRS_navigation 
   Caption         =   "[Run Simple edition]"
   ClientHeight    =   3036
   ClientLeft      =   132
   ClientTop       =   420
   ClientWidth     =   9588
   OleObjectBlob   =   "frmRS_navigation.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmRS_navigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the navigation panel
'   UserForm frmRS_navigation

'Preparation of the pop-up
Private Sub UserForm_Initialize()
    With Me
        .caption = GetText(g_arr_Text, "USERFORM005")
        .Height = 124
        .width = 335
    End With
    'Command buttons
    With cmdNavigation
        .caption = GetText(g_arr_Text, "NAVI001")
        .WordWrap = True
    End With
    cmdNavFinishphoto.caption = GetText(g_arr_Text, "BTN004")
    cmdNavResults.caption = GetText(g_arr_Text, "BTN005")
    cmdNavWinner.caption = GetText(g_arr_Text, "BTN006")
    With cmdNavBets
        .caption = GetText(g_arr_Text, "BTN007")
        .Visible = False
    End With
End Sub

'Execution of this procedure each time the pop-up is called
Private Sub UserForm_Activate()
    With Me 'Place the navigation panel in the upper left corner
        .top = Application.ActiveWindow.top + 20
        .left = Application.ActiveWindow.left + 20
    End With
    
    'Top position and height of a button depending of the number of buttons shown
    Dim nav As Integer
    If objOption.BET_PLACED Then
        nav = 24
        cmdNavBets.Visible = True 'In case of bets are placed
    Else
        nav = 32
        cmdNavBets.Visible = False 'If no bets are placed
    End If
    
    'Adjust the buttons
    With cmdNavFinishphoto
        .top = nav * 0
        .Height = nav
    End With
    With cmdNavResults
        .top = nav * 1
        .Height = nav
    End With
    With cmdNavWinner
        .top = nav * 2
        .Height = nav
    End With
    With cmdNavBets
        .top = nav * 3
        .Height = nav
    End With
End Sub

'Click events
'------------
Private Sub cmdNavigation_Click()
    Unload Me
    Call basAuxiliary.ActivateRaceSheet 'Activate the GALOPPSIM worksheet
    Call basAuxiliary.Scroll(1, 1) 'Scroll to the upper left
End Sub

Private Sub cmdNavFinishphoto_Click()
    Call ShowFinishPhoto
    UserForm_Activate
End Sub

Private Sub cmdNavResults_Click()
    Call ShowRankingList
    UserForm_Activate
End Sub

Private Sub cmdNavWinner_Click()
    Call ShowWinnerPhoto
    UserForm_Activate
End Sub

Private Sub cmdNavBets_Click()
    If frmBettingAnalysis.Visible Then
        frmBettingAnalysis.hide
    Else
        Call ShowBets
        UserForm_Activate
    End If
End Sub
