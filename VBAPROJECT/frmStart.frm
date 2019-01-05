VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStart 
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11115
   OleObjectBlob   =   "frmStart.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This is the code for the UserForm when a new race is started

Dim i As Integer

Private Sub UserForm_Initialize()

    'Hide or show the "Track specefic settings" button
    cmdS4.Visible = False
    Select Case objRace.TRACK_SURFACE
        Case "M"
            With cmdS4
                .Visible = True
                .Caption = GetText(g_arr_Text, "TRACKSPEC001")
            End With
    End Select

    'Hide the section if the betting mode is not enabled or betting is not allowed
    If objOption.BET_MODE = False Or objRace.BETTING_ALLOWED = "N" Then
        With Me
            .lblS8.Visible = False
            .cmdS1.Visible = False
            .cmdS6.Visible = False
            .lstBetSlips.Visible = False
        End With
    End If
    
    'Format and hide the area with the bet slips as long as there is no bet placed
    With lstBetSlips
        .Font = "Courier New"
        .Height = 76
        .Visible = False
    End With
    
    'Show the "Focused run" dropdown if activated
    With Me
       .lblFocus.Visible = objOption.FOCUSED_RUN 'Label visible = TRUE or FALSE
       'Horse preview
       .lblFocusColour.Visible = objOption.FOCUSED_RUN 'Preview horse visible = TRUE or FALSE
       .lblFocusColour.Caption = "" 'no text
       .lblFocusColour.BorderStyle = fmBorderStyleSingle 'draw a border around the preview horse
       .lblFocusColour.BorderColor = &H8000000D 'blue border
       With .cboFocus
            .Clear
            .Visible = objOption.FOCUSED_RUN 'ComboBox visible = True Or False
            .ColumnCount = 2 'two column for the horse name and horse number
            .ColumnWidths = "120 Pt;0 Pt"
            .Style = fmStyleDropDownList 'allow only values from the item list, no free entries
        End With
    End With

    'Populate the "Focused run" dropdown with horses
    If objOption.FOCUSED_RUN Then
        objOption.FOCUSED_NR = 0 'reset the focused horse
        For i = 1 To g_wksRaceData.Cells(rows.count, 6).End(xlUp).row - 1 'find the last row in the STATUS column
            If g_wksRaceData.Cells(i, 6).Value = "START" Then
                With cboFocus
                    .AddItem
                    .List(.ListCount - 1, 0) = _
                        g_wksRaceData.Cells(i, 7).Value & " (#" & _
                        g_wksRaceData.Cells(i, 5).Value & ")" 'horse name and number (for displaying)
                    .List(.ListCount - 1, 1) = _
                        g_wksRaceData.Cells(i, 5).Value 'horse number (for technical usage)
                End With
            End If
        Next i
    End If
    
    With cmdS2
        .SetFocus
    End With
    
    'Display the UserForm in the center of the Window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
    
End Sub

Private Sub cmdS1_Click() 'Click on "New bet slip"
    Call Gambler
End Sub

Private Sub cmdS4_Click() 'track specific settings
    Select Case objRace.TRACK_SURFACE
        Case "M"
            With frmTrackSettingsMudflats
                .Caption = objRace.TRACK_SURFACE_TEXT
                .show (vbModal) 'modal
            End With
        Case Else
    
    End Select
End Sub

Private Sub lstBetSlips_Click()
    Dim i As Integer
    For i = 0 To lstBetSlips.ListCount - 1
        If lstBetSlips.SELECTED(i) Then
            Call ShowReceipt(i + 1)
            Exit For
        End If
    Next i
End Sub

Private Sub cmdS2_Click() 'Start the race
    'Check whether a horse on focus is selected
    If objOption.FOCUSED_RUN = True And objOption.FOCUSED_NR = 0 Then
        'Pop-up
            'Set the button mode
            g_strMsgButtons = "OK"
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_tool
            g_strMsgText = g_arr_Grammar(1) & " " & GetText(g_arr_Text, "START011")
            'Display the pop-up
            frmMsg_Attention.show
        Exit Sub
    End If
    
    objRace.STARTED = True
    If g_colBetSlips.count = 0 Then objOption.BET_PLACED = False
    
    Unload Me
End Sub

Private Sub cmdS5_Click()
    Call ShowSpeed(False)
End Sub

Private Sub cmdS6_Click()
    Call ShowSpeed(True)
End Sub

'Change focused horse
Private Sub cboFocus_Change()
    For i = 1 To UBound(g_arr_varHorses)
        If CInt(cboFocus.List(cboFocus.ListIndex, 1)) = g_arr_varHorses(i, 11) Then 'CInt ---> Compare the same data types (Integer)
            objOption.FOCUSED_NR = g_arr_varHorses(i, 11) 'Set number of the focused horse
            
            'Display horse colour
            If IsArray(g_arr_varHorses(i, 2)) Then
                lblFocusColour.BackColor = g_arr_varHorses(i, 2)(7) 'Multicoloured horse: Take the head
            Else
                lblFocusColour.BackColor = g_arr_varHorses(i, 2) 'Monochrome horse
            End If
            
            Exit For 'Leave For Loop
        End If
    Next i
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then 'UserForm closed
        'Set the button mode
        g_strMsgButtons = "YesNo"
        'Assign the text for the pop-up
        g_strMsgCaption = g_c_tool
        g_strMsgText = GetText(g_arr_Text, "START004")
        'Display the pop-up
        frmMsg_MultiPurpose.show
        'Evaluate return value
        If g_strButtonPressed = "YES" Then '>>Yes<< clicked
            objRace.STARTED = False
            If g_strPlayMode = "RS" Then Call RS_ShowNavigationPanel(False)
            If g_strPlayMode = "AI" Then Call AI_ExcelModeEnd
        Else '>>No<< clicked
            Cancel = 1
        End If
    Else 'Start button pressed
        objRace.STARTED = True
    End If
End Sub
