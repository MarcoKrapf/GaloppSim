Attribute VB_Name = "basInputOutput"
Option Explicit
Option Private Module

'This module contains procedures for saving and loading files
'   Module basInputOutput

'Save all settings from the "Race options" pop-up to a text file on the hard disk
Public Sub SaveRaceOptions()

    Dim strFile As String 'File name and path
    Dim intFileNr As Integer 'Output channel
    
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
        strFile = g_defaultPath & g_c_defaultRaceOptionsFile & g_c_defaultFileType
        intFileNr = FreeFile 'Assign the next free number
        
        Open strFile For Output As #intFileNr 'Open the output channel
            With frmOptionsRace
                'Write the settings into a text file
                'Boolean values (like Checkboxes or Option Buttons) need to be converted to integer values (False >> 0 // True >> -1)
                Print #intFileNr, CInt(.optRS_ro01.Value) 'objOption.TACTICS (FALSE)
                Print #intFileNr, CInt(.optRS_ro02.Value) 'objOption.TACTICS (TRUE)
                Print #intFileNr, CInt(.chkRS_ro18.Value) 'objOption.MOMENTUM
                Print #intFileNr, CInt(.chkRS_ro12a.Value) 'objOption.SLIPSTREAM
                Print #intFileNr, CInt(.chkRS_ro12b.Value) 'objOption.SLIPSTREAM_DBL
                Print #intFileNr, CInt(.chkRS_ro12c.Value) 'objOption.SLIPSTREAM_SHOW
                Print #intFileNr, CInt(.opt_foc01.Value) 'objOption.FOCUSED_RUN (STANDARD)
                Print #intFileNr, CInt(.opt_foc02.Value) 'objOption.FOCUSED_RUN (FOCUS_HORSE)
                Print #intFileNr, CInt(.opt_foc03.Value) 'objOption.FOCUSED_RUN (FOCUS_LEADER)
                Print #intFileNr, CInt(.chk_foc02b.Value) 'objOption.HIGHLIGHT_FOC
                Print #intFileNr, CInt(.chkRS_ro11a.Value) 'objOption.BET_MODE
                Print #intFileNr, CInt(.chkRS_ro11b.Value) 'objOption.BET_ANALYSIS
                Print #intFileNr, CInt(.chkRS_ro10.Value) 'objOption.RACE_INFO
                Print #intFileNr, CInt(.chkRS_ro13.Value) 'objOption.RACE_INFO_LEADER
                Print #intFileNr, CInt(.chkRS_ro14.Value) 'objOption.RACE_INFO_PROGRESS
                Print #intFileNr, .lblRS_ro07.BackColor 'objOption.RACE_INFO_COL_B
                Print #intFileNr, .lblRS_ro07.ForeColor 'objOption.RACE_INFO_COL_F
                Print #intFileNr, CInt(.chkRS_ro02.Value) 'objOption.HOOFPRINTS
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
                Print #intFileNr, CInt(.togRS_ro01.Value) 'objOption.PHOTO_BW
                Print #intFileNr, CInt(.chkTribunes.Value) 'objOption.TRIBUNES
                Print #intFileNr, CInt(.scrSpec.Value) 'objOption.SPECTATORS
                Print #intFileNr, .scrMomRefr.Value 'objOption.MOMENTUM_REFRESHRATE
                Print #intFileNr, .scrRefusalRate.Value 'objOption.REFUSAL_RATE
                Print #intFileNr, CInt(.chk_TEO1.Value) 'objOption.TACTICS_REVEAL_TAC
                Print #intFileNr, CInt(.chk_TEO2.Value) 'objOption.TACTICS_REVEAL_CURR
                Print #intFileNr, CInt(.chkRS_ro19.Value) 'objOption.AUTOFIT
            End With
        Close #intFileNr 'Close the output channel

        'Show a pop-up: "Saved successfully."
    Call ShowMessagePopup(GetText(g_arr_Text, "BTN001") & " - " & GetText(g_arr_Text, "RACEOPT037"), _
        GetText(g_arr_Text, "SUCCESS001") & vbNewLine & vbNewLine & strFile, _
        enumButton.OK, vbModal)
                
    Exit Sub
ERRORHANDLING:
    'Show a pop-up: "Error while saving."
    Call ShowMessagePopup(GetText(g_arr_Text, "BTN001") & " - " & GetText(g_arr_Text, "RACEOPT037"), _
        GetText(g_arr_Text, "ERROR005"), _
        enumButton.OK, vbModal)
End Sub

'Restore all settings from the text file on the hard disk to the "Race options" pop-up
Public Sub LoadRaceOptions()

    Dim strFile As String 'File name and path
    Dim intFileNr As Integer 'Input channel
    Dim arr_strSettings(1 To 39) As String 'String array for the values
    Dim i As Integer
    
    On Error Resume Next 'Ignore errors
    
    strFile = g_defaultPath & g_c_defaultRaceOptionsFile & g_c_defaultFileType
    
    If Dir(strFile) <> "" Then 'Check whether the file exists
    
        intFileNr = FreeFile 'Assign the next free number
        
        'Load the settings and write them into an Array
        Open strFile For Input As #intFileNr 'Open the input channel
            For i = 1 To UBound(arr_strSettings)
                Line Input #intFileNr, arr_strSettings(i)
            Next
        Close #intFileNr 'Close the input channel
        
        'Write the values from the Array into the "Race options" UserForm
        'Boolean values must be converted back from integer values (0 >> False // -1 >> True)
            With frmOptionsRace
                .optRS_ro01.Value = CBool(arr_strSettings(1)) 'objOption.TACTICS (FALSE)
                .optRS_ro02.Value = CBool(arr_strSettings(2)) 'objOption.TACTICS (TRUE)
                .chkRS_ro18.Value = CBool(arr_strSettings(3)) 'objOption.MOMENTUM
                .chkRS_ro12a.Value = CBool(arr_strSettings(4)) 'objOption.SLIPSTREAM
                .chkRS_ro12b.Value = CBool(arr_strSettings(5)) 'objOption.SLIPSTREAM_DBL
                .chkRS_ro12c.Value = CBool(arr_strSettings(6)) 'objOption.SLIPSTREAM_SHOW
                .opt_foc01.Value = CBool(arr_strSettings(7)) 'objOption.FOCUSED_RUN (STANDARD)
                .opt_foc02.Value = CBool(arr_strSettings(8)) 'objOption.FOCUSED_RUN (FOCUS_HORSE)
                .opt_foc03.Value = CBool(arr_strSettings(9)) 'objOption.FOCUSED_RUN (FOCUS_LEADER)
                .chk_foc02b.Value = CBool(arr_strSettings(10)) 'objOption.HIGHLIGHT_FOC
                .chkRS_ro11a.Value = CBool(arr_strSettings(11)) 'objOption.BET_MODE
                .chkRS_ro11b.Value = CBool(arr_strSettings(12)) 'objOption.BET_ANALYSIS
                .chkRS_ro10.Value = CBool(arr_strSettings(13)) 'objOption.RACE_INFO
                .chkRS_ro13.Value = CBool(arr_strSettings(14)) 'objOption.RACE_INFO_LEADER
                .chkRS_ro14.Value = CBool(arr_strSettings(15)) 'objOption.RACE_INFO_PROGRESS
                .lblRS_ro07.BackColor = arr_strSettings(16) 'objOption.RACE_INFO_COL_B
                .lblRS_ro07.ForeColor = arr_strSettings(17) 'objOption.RACE_INFO_COL_F
                .chkRS_ro02.Value = CBool(arr_strSettings(18)) 'objOption.HOOFPRINTS
                .scrRS_ro02.Value = arr_strSettings(19) 'objOption.METRES_DISPLAY
                .chkRS_ro04.Value = CBool(arr_strSettings(20)) 'objOption.NAMES_LEFT
                .chkRS_ro05.Value = CBool(arr_strSettings(21)) 'objOption.COLOURS_LEFT
                .chkRS_ro06.Value = CBool(arr_strSettings(22)) 'objOption.HIGHLIGHT_FAV
                .chkRS_ro07.Value = CBool(arr_strSettings(23)) 'objOption.NAMES_FINISH
                .chkRS_ro08.Value = CBool(arr_strSettings(24)) 'objOption.RANKING_COL
                .chkRS_ro09.Value = CBool(arr_strSettings(25)) 'objOption.RANKING_DELAY
                .opt_ro01.Value = CBool(arr_strSettings(26)) 'objOption.RACE_INFO_POP
                .opt_ro02.Value = CBool(arr_strSettings(27)) 'objOption.RACE_INFO_WKS
                .scrRS_ro03.Value = arr_strSettings(28) 'objOption.SPEED_FACTOR
                .chkRS_ro15.Value = CBool(arr_strSettings(29)) 'objOption.REFUSE_RUN
                .chkRS_ro16.Value = CBool(arr_strSettings(30)) 'objOption.SPEECH
                .chkRS_ro17.Value = CBool(arr_strSettings(31)) 'objOption.NAMES_PHOTO
                .togRS_ro01.Value = CBool(arr_strSettings(32)) 'objOption.PHOTO_BW
                .chkTribunes.Value = CBool(arr_strSettings(33)) 'objOption.TRIBUNES
                .scrSpec.Value = arr_strSettings(34) 'objOption.SPECTATORS
                .scrMomRefr.Value = arr_strSettings(35) 'objOption.MOMENTUM_REFRESHRATE
                .scrRefusalRate.Value = arr_strSettings(36) 'objOption.REFUSAL_RATE
                .chk_TEO1.Value = CBool(arr_strSettings(37)) 'objOption.TACTICS_REVEAL_TAC
                .chk_TEO2.Value = CBool(arr_strSettings(38)) 'objOption.TACTICS_REVEAL_CURR
                .chkRS_ro19.Value = CBool(arr_strSettings(39)) 'objOption.AUTOFIT
            End With
        'Show a pop-up: "Restored successfully."
        Call ShowMessagePopup(GetText(g_arr_Text, "BTN001") & " - " & GetText(g_arr_Text, "RACEOPT038"), _
            GetText(g_arr_Text, "ERROR008"), _
            enumButton.OK, vbModal)
    Else
        'Show a pop-up: "File not found."
        Call ShowMessagePopup(GetText(g_arr_Text, "BTN001") & " - " & GetText(g_arr_Text, "RACEOPT038"), _
            GetText(g_arr_Text, "ERROR006"), _
            enumButton.OK, vbModal)
    End If
End Sub

'Add an entry to the error log file on the hard disk
'The optional parameter can be submitted by the calling procedure for logging extra information but is not mandatory
Public Sub WriteErrorLog(timestamp As Date, error As ErrObject, cmod As Object, procedure As String, Optional extrainfo As String)
    Dim intFileNr As Integer 'Output channel
    intFileNr = FreeFile 'Assign the next free number

    Open g_errorLogPath & Application.PathSeparator & g_c_errorLogFileName & ".txt" For Append As #intFileNr
        Print #intFileNr, timestamp & "|" & ThisWorkbook.name & "|" & cmod & "|" & procedure & "|Error " & error.Number & " " & error.Description & "|" & extrainfo
    Close #intFileNr
    
End Sub
