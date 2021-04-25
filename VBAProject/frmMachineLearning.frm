VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMachineLearning 
   Caption         =   "Machine Learning Race Simulation"
   ClientHeight    =   6852
   ClientLeft      =   -1428
   ClientTop       =   -5808
   ClientWidth     =   5568
   OleObjectBlob   =   "frmMachineLearning.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmMachineLearning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with the GUI for Machine Learning races
'   UserForm frmMachineLearning

Dim intMLnumberOfRaces As Integer 'Numer of selected races
Dim lngMLrepetitions As Long 'Number of repetitions for each race
Dim varSave As Variant 'Folder and file for saving

Private Sub btnMLStart_Click()

    Dim i As Integer
    Dim colMLRaces As Collection
    Set colMLRaces = New Collection

    With listboxMLraces
        For i = 0 To .ListCount - 1
            If .SELECTED(i) Then
                colMLRaces.Add (.List(i))
            End If
        Next i
    End With

    Call basDeveloperTools.MachineLearningSimulation(colMLRaces, lngMLrepetitions, varSave)

End Sub

Private Sub listboxMLraces_Change()
    Dim item As Integer
    Dim count As Integer
    
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    With listboxMLraces
        For item = 0 To .ListCount - 1
            If .SELECTED(item) Then
                count = count + 1
            End If
        Next item
    End With
    
    intMLnumberOfRaces = count
    
    lblScopeSelection = "Race selection (" & intMLnumberOfRaces & ")"
    lblCountMLexecutions.caption = "= " & Format((lngMLrepetitions * intMLnumberOfRaces), "#,###,###,##0") & " races to simulate"
    If intMLnumberOfRaces * lngMLrepetitions > 2147486647# Then Err.Raise (1)
    
    btnMLStart.Enabled = True
    btnMLStart.BackColor = &H80FF80
    
    Exit Sub
ERRORHANDLING:
    MsgBox "ML Overkill", vbExclamation, ""
    textboxMLrepetitions.Value = 1
    Resume
End Sub

Private Sub textboxMLrepetitions_Change()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    If Not IsNumeric(textboxMLrepetitions.Value) Then textboxMLrepetitions.Value = 1
    lngMLrepetitions = textboxMLrepetitions.Value
    
    If intMLnumberOfRaces * lngMLrepetitions > 2147486647# Then Err.Raise (1)
    
    lblCountMLexecutions.caption = "= " & Format((lngMLrepetitions * intMLnumberOfRaces), "#,###,###,##0") & " race(s) to simulate"

    Exit Sub
ERRORHANDLING:
    MsgBox "ML Overkill", vbExclamation, ""
    textboxMLrepetitions.Value = 1
    Resume
End Sub

'Select the folder for the CSV export file
Private Sub btnDataFile_Click()
    
    varSave = Application.GetSaveAsFilename( _
    InitialFileName:=g_MLdataFileName, _
    FileFilter:="CSV File (semicolon separated) (*.csv), *.csv", _
    Title:="Machine Learning Simulation Export")
    
    If varSave = False Then varSave = g_defaultMLpath & Application.PathSeparator & g_MLdataFileName & ".csv"

    frmMachineLearning.chkExport.caption = "Export data to " & vbNewLine & varSave
    btnDataFile.BackColor = &H8000000F
    
End Sub

Private Sub UserForm_Initialize()
    Dim item As Integer
    
    With Me
        .width = 385
        .Height = 425
    End With
    
    'Fill the list box with all installed races
    With listboxMLraces
        .MultiSelect = fmMultiSelectExtended
        For item = 1 To g_colRacesInstalled.count
            .AddItem g_colRacesInstalled(item)
        Next item
    End With
    
    textboxMLrepetitions.Value = 1
    chkExport.Value = True
    chkDebug.Value = False
    btnDataFile.BackColor = &H80FFFF
    btnMLStart.BackColor = &H8000000F
    btnMLStart.Enabled = False
    
    varSave = g_defaultMLpath & Application.PathSeparator & g_MLdataFileName & ".csv"
    frmMachineLearning.chkExport.caption = "Export data to" & vbNewLine _
            & varSave
    
End Sub
