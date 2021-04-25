VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSuperDev 
   Caption         =   "Super Developer Tools"
   ClientHeight    =   6276
   ClientLeft      =   -84
   ClientTop       =   -444
   ClientWidth     =   8724
   OleObjectBlob   =   "frmSuperDev.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmSuperDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the Super Admin GUI
'   UserForm frmSuperDev

'Call the Test Suite GUI
Private Sub btnTestSuite_Click()
    frmTestSuite.show (vbModeless)
    Unload Me 'Close the Super Admin GUI
End Sub

'Call the Machine Learning Race Simulation GUI
Private Sub btnMLsimulation_Click()
    frmMachineLearning.show (vbModeless)
    Unload Me 'Close the Super Admin GUI
End Sub

Private Sub UserForm_Initialize()
    With Me
        .width = 455
        .Height = 300
    End With
    
    'Race Speed Equilizer
    With scrEq1
        .min = 0
        .max = 1
        .Value = g_arr_Developer(1)
    End With
    With scrEq2
        .min = 0
        .max = 1
        .Value = g_arr_Developer(2)
    End With
    With scrEq3
        .min = 0
        .max = 1
        .Value = g_arr_Developer(3)
    End With
    
    'Toggle buttons
    With togDev1
        .Value = g_skipDelay
        If .Value = True Then
            .BackColor = &H80FF80 'Green
        Else
            .BackColor = &H8080FF 'Red
        End If
    End With
    With togDev2
        .Value = g_errorLogging
        If .Value = True Then
            .BackColor = &H80FF80 'Green
        Else
            .BackColor = &H8080FF 'Red
        End If
    End With
    With togDev3
        .Value = g_payoutLogging
        If .Value = True Then
            .BackColor = &H80FF80 'Green
        Else
            .BackColor = &H8080FF 'Red
        End If
    End With
    
    'Display the UserForm in the center of the Window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
    
End Sub

'Slider for the basic speed factor
Private Sub scrEq1_Change()
    g_arr_Developer(1) = scrEq1.Value
End Sub

'Slider for the form factor
Private Sub scrEq2_Change()
    g_arr_Developer(2) = scrEq2.Value
End Sub

'Slider for the loop factor
Private Sub scrEq3_Change()
    g_arr_Developer(3) = scrEq3.Value
End Sub

'Toggle button for skipping delay commands
Private Sub togDev1_Click()
    With togDev1
        If .Value = True Then
            .BackColor = &H80FF80 'Green
        Else
            .BackColor = &H8080FF 'Red
        End If
        g_skipDelay = .Value
    End With
End Sub

'Toggle button for error logging
Private Sub togDev2_Click()
    With togDev2
        If .Value = True Then
            .BackColor = &H80FF80 'Green
        Else
            .BackColor = &H8080FF 'Red
        End If
        g_errorLogging = .Value
    End With
End Sub

'Toggle button for skipping payout logging
Private Sub togDev3_Click()
    With togDev3
        If .Value = True Then
            .BackColor = &H80FF80 'Green
        Else
            .BackColor = &H8080FF 'Red
        End If
        g_payoutLogging = .Value
    End With
End Sub
