VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestSuite 
   Caption         =   "RACE TEST AUTOMATION (RTA)"
   ClientHeight    =   6288
   ClientLeft      =   -1212
   ClientTop       =   -4908
   ClientWidth     =   7080
   OleObjectBlob   =   "frmTestSuite.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmTestSuite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the Test Suite for automatic testing
'   UserForm frmTestSuite

Private Sub btnTestStart_Click()
    
    If multiPage.Value = 0 Then 'Standard

        Dim colTestRaces As Collection
        Dim i As Integer
        Set colTestRaces = New Collection
    
        With listboxRaces_std
            For i = 0 To .ListCount - 1
                If .SELECTED(i) Then
                    colTestRaces.Add (.List(i))
                End If
            Next i
        End With
    
        Call TestStart_std(colTestRaces, opt2_std.Value, chk1_std.Value, _
                chk2_std.Value, chk3_std.Value, chk4_std.Value, chk6_std.Value, chk7_std.Value)
    
    Else 'Test repository (not yet implemented)

    End If
End Sub

Private Sub listboxRaces_std_Change()
    Dim item As Integer
    Dim count As Integer
    
    With listboxRaces_std
        For item = 0 To .ListCount - 1
            If .SELECTED(item) Then
                count = count + 1
            End If
        Next item
    End With
    
    lblScopeSelection_std = "Test scope selection (" & count & ")"
    btnTestStart.Enabled = True
End Sub

Private Sub listboxRaces_rep_Change()
    Dim item As Integer
    Dim count As Integer
    
    With listboxRaces_rep
        For item = 0 To .ListCount - 1
            If .SELECTED(item) Then
                count = count + 1
            End If
        Next item
    End With
    
    lblScopeSelection_rep = "Test scope selection (" & count & ")"
    btnTestStart.Enabled = True
End Sub

Private Sub opt1_std_Click()
    chk1_std.Enabled = False
    chk2_std.Enabled = False
    chk3_std.Enabled = False
    chk4_std.Enabled = False
    chk6_std.Enabled = False
    chk7_std.Enabled = False
End Sub

Private Sub opt2_std_Click()
    chk1_std.Enabled = True
    chk2_std.Enabled = True
    chk3_std.Enabled = True
    chk4_std.Enabled = True
    chk6_std.Enabled = True
    chk7_std.Enabled = True
End Sub

Private Sub togMinMax_Click()
    With Me
        If togMinMax.Value = True Then
            .width = 100
            .Height = 46
        Else
            .width = 500
            .Height = 420
        End If
    End With
End Sub



Private Sub UserForm_Initialize()
    Dim item As Integer
    
    With Me
        .width = 385
        .Height = 420
    End With
    
    'Fill the "Standard" list box with all installed races
    With listboxRaces_std
        .MultiSelect = fmMultiSelectExtended
        For item = 1 To g_colRacesInstalled.count
            .AddItem g_colRacesInstalled(item)
        Next item
    End With
    
    'Fill the "Test repository" list box with the test cases from Worksheet "TESTCASES"
    With listboxRaces_rep
        .MultiSelect = fmMultiSelectExtended
        For item = 4 To g_wksTCASE.Cells(1, Columns.count).End(xlToLeft).Column
            .AddItem g_wksTCASE.Cells(1, item).Value _
            & " (" & g_wksTCASE.Cells(2, item).Value & ")"
        Next item
    End With
    
    btnTestStart.Enabled = False
    
    multiPage.Value = 0 'Select the first page
    multiPage.Pages(1).Enabled = False
End Sub
