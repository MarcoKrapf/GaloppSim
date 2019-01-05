Attribute VB_Name = "basInputOutput"
Option Explicit
Option Private Module

'This module contains procedures for saving and loading files

Public Sub SaveRaceOptions()

    Dim strFile As String 'File name and path
    Dim intFileNr As Integer 'Output channel
    
    On Error GoTo ERRORHANDLING
    
        strFile = g_defaultPath & g_c_defaultRaceOptionsFile & g_c_defaultFileType
        intFileNr = FreeFile 'Assign the next free number
        
        Open strFile For Output As #intFileNr 'Open the output channel
            With frmOptionsRace
                'Write settings into a text file
                Print #intFileNr, CInt(.optRS_ro01.Value) 'objOption.TACTICS 0
                Print #intFileNr, CInt(.optRS_ro02.Value) 'objOption.TACTICS 3
                Print #intFileNr, CInt(.optRS_ro03.Value) 'objOption.TACTICS 6
                Print #intFileNr, CInt(.chkRS_ro12a.Value) 'objOption.SLIPSTREAM
                Print #intFileNr, CInt(.chkRS_ro12b.Value) 'objOption.SLIPSTREAM_DBL
                Print #intFileNr, CInt(.chkRS_ro12c.Value) 'objOption.SLIPSTREAM_SHOW
                Print #intFileNr, CInt(.chkRS_ro01a.Value) 'objOption.FOCUSED_RUN
                Print #intFileNr, CInt(.chkRS_ro01b.Value) 'objOption.HIGHLIGHT_FOC
                Print #intFileNr, CInt(.chkRS_ro11a.Value) 'objOption.BET_MODE
                Print #intFileNr, CInt(.chkRS_ro11b.Value) 'objOption.BET_ANALYSIS
                Print #intFileNr, CInt(.chkRS_ro10.Value) 'objOption.RACE_INFO
                Print #intFileNr, CInt(.chkRS_ro13.Value) 'objOption.RACE_INFO_LEADER
                Print #intFileNr, CInt(.chkRS_ro14.Value) 'objOption.RACE_INFO_PROGRESS
                Print #intFileNr, .lblRS_ro07.BackColor 'objOption.RACE_INFO_COL_B
                Print #intFileNr, .lblRS_ro07.ForeColor 'objOption.RACE_INFO_COL_F
                Print #intFileNr, CInt(.chkRS_ro02.Value) 'objOption.HOOFPRINTS
                Print #intFileNr, .scrRS_ro01.Value 'objOption.ZOOM_LEVEL
                Print #intFileNr, .scrRS_ro02.Value 'objOption.METRES_DISPLAY
                Print #intFileNr, CInt(.chkRS_ro04.Value) 'objOption.NAMES_LEFT
                Print #intFileNr, CInt(.chkRS_ro05.Value) 'objOption.COLOURS_LEFT
                Print #intFileNr, CInt(.chkRS_ro06.Value) 'objOption.HIGHLIGHT_FAV
                Print #intFileNr, CInt(.chkRS_ro07.Value) 'objOption.NAMES_FINISH
                Print #intFileNr, CInt(.chkRS_ro08.Value) 'objOption.RANKING_COL
                Print #intFileNr, CInt(.chkRS_ro09.Value) 'objOption.RANKING_DELAY
                Print #intFileNr, CInt(.opt_ro01.Value) 'objOption.RACE_INFO_POP
                Print #intFileNr, CInt(.opt_ro02.Value) 'objOption.RACE_INFO_WKS
                Print #intFileNr, .scrRS_ro03.Value 'objOption.SPEED_FACTOR
                Print #intFileNr, CInt(.chkRS_ro15.Value) 'objOption.REFUSE_RUN
                Print #intFileNr, CInt(.chkRS_ro16.Value) 'objOption.SPEECH
                Print #intFileNr, CInt(.chkRS_ro17.Value) 'objOption.NAMES_PHOTO
            End With
                Print #intFileNr, objOption.LUGWORMS
                Print #intFileNr, objOption.TIDE
            Print #intFileNr, CInt(frmOptionsRace.togRS_ro01.Value) 'objOption.PHOTO_BW
        Close #intFileNr 'Close the output channel

        'Show a pop-up
            'Set the button mode
            g_strMsgButtons = "OK"
            'Assign the text for the pop-up
            g_strMsgCaption = GetText(g_arr_Text, "BTN001") & " - " _
                & GetText(g_arr_Text, "RACEOPT037")
            g_strMsgText = GetText(g_arr_Text, "SUCCESS001") _
                & vbNewLine & vbNewLine & strFile
            'Display the pop-up (modal)
            frmMsg_MultiPurpose.show (vbModal)
                
    Exit Sub
    
ERRORHANDLING:
    'Show a pop-up
        'Set the button mode
        g_strMsgButtons = "OK"
        'Assign the text for the pop-up
        g_strMsgCaption = GetText(g_arr_Text, "BTN001") & " - " _
            & GetText(g_arr_Text, "RACEOPT037")
        g_strMsgText = GetText(g_arr_Text, "ERROR005")
        'Display the pop-up (modal)
        frmMsg_MultiPurpose.show (vbModal)
End Sub

Public Sub LoadRaceOptions()

    Dim strFile As String 'File name and path
    Dim intFileNr As Integer 'Output channel
    Dim arr_strSettings(1 To 33) As String 'String array for the values
    Dim i As Integer
    
    On Error Resume Next 'Ignore errors
    
    strFile = g_defaultPath & g_c_defaultRaceOptionsFile & g_c_defaultFileType
    
    If Dir(strFile) <> "" Then
    
        intFileNr = FreeFile 'Assign the next free number
        
        'Load settings and write them into an array
        Open strFile For Input As #intFileNr 'Open the input channel
            For i = 1 To UBound(arr_strSettings)
                Line Input #intFileNr, arr_strSettings(i)
            Next
        Close #intFileNr 'Close the input channel
        
        'Write the settings into the "Race options" UserForm
            With frmOptionsRace
                .optRS_ro01.Value = CBool(arr_strSettings(1)) 'objOption.TACTICS 0
                .optRS_ro02.Value = CBool(arr_strSettings(2)) 'objOption.TACTICS 3
                .optRS_ro03.Value = CBool(arr_strSettings(3)) 'objOption.TACTICS 6
                .chkRS_ro12a.Value = CBool(arr_strSettings(4)) 'objOption.SLIPSTREAM
                .chkRS_ro12b.Value = CBool(arr_strSettings(5)) 'objOption.SLIPSTREAM_DBL
                .chkRS_ro12c.Value = CBool(arr_strSettings(6)) 'objOption.SLIPSTREAM_SHOW
                .chkRS_ro01a.Value = CBool(arr_strSettings(7)) 'objOption.FOCUSED_RUN
                .chkRS_ro01b.Value = CBool(arr_strSettings(8)) 'objOption.HIGHLIGHT_FOC
                .chkRS_ro11a.Value = CBool(arr_strSettings(9)) 'objOption.BET_MODE
                .chkRS_ro11b.Value = CBool(arr_strSettings(10)) 'objOption.BET_ANALYSIS
                .chkRS_ro10.Value = CBool(arr_strSettings(11)) 'objOption.RACE_INFO
                .chkRS_ro13.Value = CBool(arr_strSettings(12)) 'objOption.RACE_INFO_LEADER
                .chkRS_ro14.Value = CBool(arr_strSettings(13)) 'objOption.RACE_INFO_PROGRESS
                .lblRS_ro07.BackColor = arr_strSettings(14) 'objOption.RACE_INFO_COL_B
                .lblRS_ro07.ForeColor = arr_strSettings(15) 'objOption.RACE_INFO_COL_F
                .chkRS_ro02.Value = CBool(arr_strSettings(16)) 'objOption.HOOFPRINTS
                .scrRS_ro01.Value = arr_strSettings(17) 'objOption.ZOOM_LEVEL
                .scrRS_ro02.Value = arr_strSettings(18) 'objOption.METRES_DISPLAY
                .chkRS_ro04.Value = CBool(arr_strSettings(19)) 'objOption.NAMES_LEFT
                .chkRS_ro05.Value = CBool(arr_strSettings(20)) 'objOption.COLOURS_LEFT
                .chkRS_ro06.Value = CBool(arr_strSettings(21)) 'objOption.HIGHLIGHT_FAV
                .chkRS_ro07.Value = CBool(arr_strSettings(22)) 'objOption.NAMES_FINISH
                .chkRS_ro08.Value = CBool(arr_strSettings(23)) 'objOption.RANKING_COL
                .chkRS_ro09.Value = CBool(arr_strSettings(24)) 'objOption.RANKING_DELAY
                .opt_ro01.Value = CBool(arr_strSettings(25)) 'objOption.RACE_INFO_POP
                .opt_ro02.Value = CBool(arr_strSettings(26)) 'objOption.RACE_INFO_WKS
                .scrRS_ro03.Value = arr_strSettings(27) 'objOption.SPEED_FACTOR
                .chkRS_ro15.Value = CBool(arr_strSettings(28)) 'objOption.REFUSE_RUN
                .chkRS_ro16.Value = CBool(arr_strSettings(29)) 'objOption.SPEECH
                .chkRS_ro17.Value = CBool(arr_strSettings(30)) 'objOption.NAMES_PHOTO
                .togRS_ro01.Value = CBool(arr_strSettings(33)) 'objOption.PHOTO_BW
            End With
                objOption.LUGWORMS = arr_strSettings(31)
                objOption.TIDE = arr_strSettings(32)
    Else
        'Show a pop-up
            'Set the button mode
            g_strMsgButtons = "OK"
            'Assign the text for the pop-up
            g_strMsgCaption = GetText(g_arr_Text, "BTN001") & " - " _
                & GetText(g_arr_Text, "RACEOPT038")
            g_strMsgText = GetText(g_arr_Text, "ERROR006")
            'Display the pop-up (modal)
            frmMsg_MultiPurpose.show (vbModal)
    End If
    
    On Error GoTo 0
End Sub
