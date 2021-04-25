VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReplayOptions 
   Caption         =   "[caption]"
   ClientHeight    =   4596
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8640
   OleObjectBlob   =   "frmReplayOptions.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmReplayOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with options when a race replay is started
'   UserForm frmReplayOptions

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    'UserForm settings
    '-----------------
    If objOption.SPEEDMONITOR Then i = i + 10 * UBound(g_arr_varHorses())
    If objOption.FOCUSED_RUN = enumCamera.focus_horse Then i = i + 20
    If objOption.SPEEDMONITOR And objOption.FOCUSED_RUN Then i = i + 30
    
    With btnReplayPopupOK
        .left = 250
        .width = 60
        .top = i + 40
        .caption = GetText(g_arr_Text, "BTN014")
    End With
    
    With Me
        .caption = GetText(g_arr_Text, "BTN025")
        .width = 340
        .Height = i + 100
    End With
    
    'Race Speed Monitor section
    '--------------------------
    If objOption.SPEEDMONITOR Then
    
        'Create a label for the section caption
        Set g_objLabel = frmReplayOptions.Controls.Add("Forms.Label.1", , True)
        With g_objLabel
            With .Font
                .name = "Tahoma"
                .size = 12
                .Bold = True
            End With
            .left = 10
            .top = 10
            .width = 300
            .TextAlign = fmTextAlignLeft
            .caption = GetText(g_arr_Text, "START016")
        End With
        
        'Create a ListBox for the horse selection
        Set g_objListbox = frmReplayOptions.Controls.Add("Forms.ListBox.1", , True)
        With g_objListbox
            .left = 10
            .top = 30
            .width = 300
            .Height = 10 * UBound(g_arr_varHorses())
            .ColumnCount = 2 'Provide two columns for the horse name and number
            .ColumnWidths = "295 Pt;0 Pt" '300 pixels for the name, 0 for the number
        End With

        'Populate the ListBox with horses
        For i = 1 To UBound(g_arr_varHorses())
            If g_arr_varHorses(i, 0) = "START" Then 'Consider only horses that are starting
                With g_objListbox 'ListBox with the horses
                    .AddItem
                    .List(.ListCount - 1, 0) = _
                        g_arr_varHorses(i, 1) & " (#" & _
                        g_arr_varHorses(i, 11) & ")" 'Column 1 (visible): Compile the entry with the horse name and number (for displaying)
                    .List(.ListCount - 1, 1) = _
                        g_arr_varHorses(i, 11) 'Column 2 (invisible): Horse number (for technical usage)
                End With
            End If
        Next i
        g_objListbox.MultiSelect = fmMultiSelectMulti
    End If
    
    'Focused Run section
    '-------------------
    If objOption.FOCUSED_RUN = enumCamera.focus_horse Then
    
        'Create a label for the section caption
        Set g_objLabel = frmReplayOptions.Controls.Add("Forms.Label.1", , True)
        With g_objLabel
            With .Font
                .name = "Tahoma"
                .size = 12
                .Bold = True
            End With
            .left = 10
            If objOption.SPEEDMONITOR Then
                .top = 40 + 10 * UBound(g_arr_varHorses())
            Else
                .top = 10
            End If
            .width = 300
            .TextAlign = fmTextAlignLeft
            .caption = g_arr_Grammar(1) & " " & GetText(g_arr_Text, "START006") & " " & GetText(g_arr_Text, "RACEOPT027")
        End With

        'Create a DropDown for the horse selection
        Set g_objDropdown = frmReplayOptions.Controls.Add("Forms.comboBox.1", , True)
        With g_objDropdown
            .left = 10
            If objOption.SPEEDMONITOR Then
                .top = 60 + 10 * UBound(g_arr_varHorses())
            Else
                .top = 30
            End If
            .width = 300
            .ColumnCount = 2 'Provide two columns for the horse name and number
            .ColumnWidths = "300 Pt;0 Pt" '300 pixels for the name, 0 for the number
            .Style = fmStyleDropDownList 'Allow only values from the item list, no free entries
        End With

        'Populate the DropDown with horses
        For i = 1 To UBound(g_arr_varHorses())
            If g_arr_varHorses(i, 0) = "START" Then 'Consider only horses that are starting
                With g_objDropdown 'DropDown with the horses
                    .AddItem
                    .List(.ListCount - 1, 0) = _
                        g_arr_varHorses(i, 1) & " (#" & _
                        g_arr_varHorses(i, 11) & ")" 'Column 1 (visible): Compile the entry with the horse name and number (for displaying)
                    .List(.ListCount - 1, 1) = _
                        g_arr_varHorses(i, 11) 'Column 2 (invisible): Horse number (for technical usage)
                End With
            End If
        Next i
        
        objOption.FOCUSED_NR = 0
    End If
    
    'Display the UserForm in the center of the Window
    Call basAuxiliary.PlaceUserFormInCenter(Me)

End Sub

Private Sub btnReplayPopupOK_Click()

    Dim checkOK As Boolean

    'Focused Run: Get the selected entry
    If objOption.FOCUSED_RUN Then
        If g_objDropdown.ListIndex = -1 Then
            Call ShowInfoPopup(g_c_tool, _
                GetText(g_arr_Text, "START017") & " " & g_arr_Grammar(2) & " " & GetText(g_arr_Text, "START011") _
                , True, vbModal, 24)
            checkOK = False
        Else
            objOption.FOCUSED_NR = g_objDropdown.List(g_objDropdown.ListIndex, 1)
            checkOK = True
        End If
    End If
    
    'Race Speed Monitor: Get the selected entries
    If objOption.SPEEDMONITOR Then
        Dim entry As Integer, count As Integer
    
        'Count the number of selected horses
        For entry = 0 To g_objListbox.ListCount - 1
            If g_objListbox.SELECTED(entry) Then count = count + 1
        Next entry
    
        'Reset the RSMon collection
        Set g_colRSMon = Nothing
        Set g_colRSMon = New Collection
        
        'Get the selected entries
        For entry = 0 To g_objListbox.ListCount - 1
            If g_objListbox.SELECTED(entry) Then g_colRSMon.Add g_objListbox.List(entry, 1)
        Next entry
        
        If g_colRSMon.count < 1 Then
            Call ShowInfoPopup(g_c_tool, GetText(g_arr_Text, "START017") & " " & g_arr_Grammar(2) _
                & " " & GetText(g_arr_Text, "START013"), True, vbModal, 24)
            checkOK = False
        ElseIf g_colRSMon.count > 3 Then
            Call ShowInfoPopup(g_c_tool, GetText(g_arr_Text, "START018") & " " & g_arr_Grammar(4) _
                & " " & GetText(g_arr_Text, "START013"), True, vbModal, 24)
            checkOK = False
        Else
            checkOK = True
        End If
    End If

    If checkOK Then Unload Me

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = 1 'Click on "X" in the upper right corner of the UserForm: don´t close the pop-up
End Sub
