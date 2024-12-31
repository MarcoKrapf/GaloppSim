VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTrackSettingsTunnels 
   Caption         =   "UserForm1"
   ClientHeight    =   4752
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6780
   OleObjectBlob   =   "frmTrackSettingsTunnels.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmTrackSettingsTunnels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the track specific settings for a tunnel race
'   UserForm frmTrackSettingsTunnels

Private Sub UserForm_Initialize()

    With Me
        .Height = 266
        .width = 356
    End With
    
    Call LabelCaptions
    
    With Me
        'Number of tunnels
        With .scrTun01
            .min = 1 'Minimum value
            .max = WorksheetFunction.max(2, objRace.METRES * 0.4 / 50) 'Maximum value
            .SmallChange = 1 'Value change when using the arrows
            .LargeChange = 3 'Value change when clicking inside the slider
            .Value = objOption.TUNNEL_COUNT
        End With

        'Total tunnel length
        With .scrTun02
            .min = 100 'Minumum value
            .max = objRace.METRES * 0.4 'Maximum value
            .SmallChange = 10 'Value change when using the arrows
            .LargeChange = 100 'Value change when clicking inside the slider
            .Value = objOption.TUNNEL_LENGTH
        End With
        
        'Road crossing the race track
        With .chkTun01
            .caption = GetText(g_arr_Text, "RACESPEC047")
            .Value = objOption.ROADCROSSING
        End With
        
        'Race is secured by the police
        With .chkTun02
            .caption = GetText(g_arr_Text, "RACESPEC048")
            .Value = objOption.POLICECAR
        End With
        
    End With

    'Display the UserForm in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
    
End Sub

'Label captions
Private Sub LabelCaptions()
    With Me
        .caption = GetText(g_arr_Text, "RACESPEC001")
        .lblTun01.caption = GetText(g_arr_Text, "RACESPEC040")
        .lblTun02.caption = GetText(g_arr_Text, "RACESPEC041")
        .lblTun03a.caption = GetText(g_arr_Text, "RACESPEC042")
        .lblTun03b.caption = GetText(g_arr_Text, "RACESPEC043")
        .lblTun04.caption = GetText(g_arr_Text, "RACESPEC044")
        .lblTun05a.caption = GetText(g_arr_Text, "RACESPEC045")
        .lblTun05b.caption = GetText(g_arr_Text, "RACESPEC046")
    End With
End Sub

'Click on the checkbox "Road crossing the race track"
Private Sub chkTun01_Click()
    If chkTun01.Value = True Then
        chkTun02.Enabled = True
    Else
        chkTun02.Value = False
        chkTun02.Enabled = False
    End If
End Sub

'Click on the OK button
Private Sub cmdTun01_Click()
    'Set the selected values
    objOption.TUNNEL_COUNT = scrTun01.Value
    objOption.TUNNEL_LENGTH = scrTun02.Value
    objOption.ROADCROSSING = chkTun01.Value
    objOption.POLICECAR = chkTun02.Value

    #If Debugging Then 'For debugging purposes
        Debug.Print
        Debug.Print "Selection:"
        Debug.Print "---------:"
        Debug.Print "Number of tunnels    : " & objOption.TUNNEL_COUNT
        Debug.Print "Tunnel length (total): " & objOption.TUNNEL_LENGTH
        Debug.Print "Tunnel length (avg)  : " & Round(objOption.TUNNEL_LENGTH / objOption.TUNNEL_COUNT, 0)
        Debug.Print "Road crossing : " & objOption.ROADCROSSING
        Debug.Print "Police car    : " & objOption.POLICECAR
        Debug.Print
    #End If
    Unload Me 'Close the pop-up
End Sub

Private Sub scrTun01_Change()
    #If Debugging Then 'For debugging purposes
        Debug.Print "Number of tunnels (min 1, max " & WorksheetFunction.max(2, objRace.METRES / 200) & ") : " & Me.scrTun01.Value
    #End If
End Sub

Private Sub scrTun02_Change()
    #If Debugging Then 'For debugging purposes
        Debug.Print "Total tunnel length (min 100, max " & objRace.METRES * 0.4 & ") : " & Me.scrTun02.Value & "m"
    #End If
End Sub
