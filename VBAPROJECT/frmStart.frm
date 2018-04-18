VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStart 
   ClientHeight    =   5235
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

'ToDo: Module description

Dim i As Integer, j As Integer

Private Sub cmdS1_Click() 'Click on "New bet slip"
    Call Gambler
End Sub

Private Sub lstBetSlips_Click()
    Dim i As Integer
    For i = 0 To lstBetSlips.ListCount - 1
        If lstBetSlips.Selected(i) Then
            Call ShowReceipt(i + 1)
            Exit For
        End If
    Next i
End Sub

Private Sub cmdS2_Click() 'Start the race
    'Check whether a horse on focus is selected
    If g_blnFocusedRun = True And g_intFocusedRun = 0 Then
        'Pop-up
            'Set the button mode
            g_strMsgButtons = "OK"
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_TOOL
            g_strMsgText = g_strTxt(97)
            'Display the pop-up
            frmMsg_Attention.Show
        Exit Sub
    End If
    
    g_blnRaceStarted = True
    Unload Me
End Sub

Private Sub cmdS6_Click()
    Call odds
End Sub

Private Sub UserForm_Initialize()

    'Hide the section if the betting mode is not enabled
    If Not g_blnBettingMode Then
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
       .lblFocus.Visible = g_blnFocusedRun 'Label visible = TRUE or FALSE
       'Horse preview
       .lblFocusColour.Visible = g_blnFocusedRun 'Preview horse visible = TRUE or FALSE
       .lblFocusColour.Caption = "" 'no text
       .lblFocusColour.Width = basMainCode.HorseSizePreview(g_byteTrackZoom)(0) 'take the first value of the array as the width
       .lblFocusColour.Height = basMainCode.HorseSizePreview(g_byteTrackZoom)(1) 'take the second value of the array as the height
       .lblFocusColour.left = 544 - basMainCode.HorseSizePreview(g_byteTrackZoom)(0) 'horizontal position of the horse according to the size
       .lblFocusColour.BorderStyle = fmBorderStyleSingle 'draw a border around the preview horse
       .lblFocusColour.BorderColor = &H8000000D 'blue border
       With .cboFocus
            .Clear
            .Visible = g_blnFocusedRun 'ComboBox visible = True Or False
            .ColumnCount = 2 'two column for the horse name and horse number
            .ColumnWidths = "120 Pt;0 Pt"
            .Style = fmStyleDropDownList 'allow only values from the item list, no free entries
        End With
    End With

    'Populate the "Focused run" dropdown with horses
    If g_blnFocusedRun Then
        g_intFocusedRun = 0 'reset the focused horse
        For i = 1 To g_wksRaceData.Cells(rows.Count, 6).End(xlUp).row - 1 'find the last row in the STATUS column
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
    Call basMainCode.PlaceUserFormInCenter(Me)
    
End Sub

'Change focused horse
Private Sub cboFocus_Change()
    For i = 1 To UBound(g_arr_varHorses)
        If CInt(cboFocus.List(cboFocus.ListIndex, 1)) = g_arr_varHorses(i, 11) Then 'CInt ---> Compare the same data types (Integer)
            g_intFocusedRun = g_arr_varHorses(i, 11) 'Set number of the focused horse
            lblFocusColour.BackColor = g_arr_varHorses(i, 2) 'Paint horse colour
            Exit For 'Leave For Loop
        End If
    Next i
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then 'UserForm closed
        'Set the button mode
        g_strMsgButtons = "YesNo"
        'Assign the text for the pop-up
        g_strMsgCaption = g_c_TOOL
        g_strMsgText = g_strTxt(103)
        'Display the pop-up
        frmMsg_MultiPurpose.Show
        'Evaluate return value
        If g_strButtonPressed = "YES" Then '>>Yes<< clicked
            g_blnRaceStarted = False
            If g_strPlayMode = "RS" Then Call RS_ShowNavigationPanel(False)
        Else '>>No<< clicked
            Cancel = 1
        End If
    Else 'Start button pressed
        g_blnRaceStarted = True
    End If
End Sub
