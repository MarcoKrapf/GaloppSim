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
                Print #intFileNr, CInt(.chk_MOM1.Value) 'objOption.MOMENTUM_BARS
                Print #intFileNr, CInt(.scrSlipstream.Value) 'objOption.SLIPSTREAM_IMPACT
                Print #intFileNr, CInt(.chk_SlipstreamShow.Value) 'objOption.SLIPSTREAM_SHOW
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
                Print #intFileNr, .scrSpec.Value 'objOption.SPECTATORS
                Print #intFileNr, .scrMomRefr.Value 'objOption.MOMENTUM_REFRESHRATE
                Print #intFileNr, .scrRefusalRate.Value 'objOption.REFUSAL_RATE
                Print #intFileNr, CInt(.chk_TEO1.Value) 'objOption.TACTICS_REVEAL_TAC
                Print #intFileNr, CInt(.chk_TEO2.Value) 'objOption.TACTICS_REVEAL_CURR
                Print #intFileNr, CInt(.chkRS_ro19.Value) 'objOption.AUTOFIT
                Print #intFileNr, CInt(.chk_MOM2.Value) 'objOption.MOMENTUM_ICONS
                Print #intFileNr, CInt(.chkRSMon.Value) 'objOption.SPEEDMONITOR
                Print #intFileNr, .scrRSMonRefr.Value 'objOption.SPEEDMON_REFRESHRATE
                Print #intFileNr, CInt(.chkFavAnn.Value) 'objOption.ANNOUNCE_FAV
                Print #intFileNr, CInt(.chkAutoSaveRace.Value) 'objOption.AUTO_SAVE
                Print #intFileNr, CInt(.optRSMon1.Value) 'objOption.RSMON_SPEED
                Print #intFileNr, CInt(.optRSMon2.Value) 'objOption.RSMON_DISTANCE
                Print #intFileNr, CInt(.opt_grid1.Value) 'objOption.STARTING_GRID_IN
                Print #intFileNr, CInt(.opt_grid2.Value) 'objOption.STARTING_GRID_BEHIND
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
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "SaveRaceOptions()")
    Close #intFileNr 'Close the output channel
End Sub

'Restore all settings from the text file on the hard disk to the "Race options" pop-up
Public Sub LoadRaceOptions()

    Dim strFile As String 'File name and path
    Dim intFileNr As Integer 'Input channel
    Dim arr_strSettings(1 To 47) As String 'String array for the values
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
                .chk_MOM1.Value = CBool(arr_strSettings(3)) 'objOption.MOMENTUM_BARS
                .scrSlipstream.Value = arr_strSettings(4) 'objOption.SLIPSTREAM_IMPACT
                .chk_SlipstreamShow.Value = CBool(arr_strSettings(5)) 'objOption.SLIPSTREAM_SHOW
                .opt_foc01.Value = CBool(arr_strSettings(6)) 'objOption.FOCUSED_RUN (STANDARD)
                .opt_foc02.Value = CBool(arr_strSettings(7)) 'objOption.FOCUSED_RUN (FOCUS_HORSE)
                .opt_foc03.Value = CBool(arr_strSettings(8)) 'objOption.FOCUSED_RUN (FOCUS_LEADER)
                .chk_foc02b.Value = CBool(arr_strSettings(9)) 'objOption.HIGHLIGHT_FOC
                .chkRS_ro11a.Value = CBool(arr_strSettings(10)) 'objOption.BET_MODE
                .chkRS_ro11b.Value = CBool(arr_strSettings(11)) 'objOption.BET_ANALYSIS
                .chkRS_ro10.Value = CBool(arr_strSettings(12)) 'objOption.RACE_INFO
                .chkRS_ro13.Value = CBool(arr_strSettings(13)) 'objOption.RACE_INFO_LEADER
                .chkRS_ro14.Value = CBool(arr_strSettings(14)) 'objOption.RACE_INFO_PROGRESS
                .lblRS_ro07.BackColor = arr_strSettings(15) 'objOption.RACE_INFO_COL_B
                .lblRS_ro07.ForeColor = arr_strSettings(16) 'objOption.RACE_INFO_COL_F
                .chkRS_ro02.Value = CBool(arr_strSettings(17)) 'objOption.HOOFPRINTS
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
                .togRS_ro01.Value = CBool(arr_strSettings(31)) 'objOption.PHOTO_BW
                .chkTribunes.Value = CBool(arr_strSettings(32)) 'objOption.TRIBUNES
                .scrSpec.Value = arr_strSettings(33) 'objOption.SPECTATORS
                .scrMomRefr.Value = arr_strSettings(34) 'objOption.MOMENTUM_REFRESHRATE
                .scrRefusalRate.Value = arr_strSettings(35) 'objOption.REFUSAL_RATE
                .chk_TEO1.Value = CBool(arr_strSettings(36)) 'objOption.TACTICS_REVEAL_TAC
                .chk_TEO2.Value = CBool(arr_strSettings(37)) 'objOption.TACTICS_REVEAL_CURR
                .chkRS_ro19.Value = CBool(arr_strSettings(38)) 'objOption.AUTOFIT
                .chk_MOM2.Value = CBool(arr_strSettings(39)) 'objOption.MOMENTUM_ICONS
                .chkRSMon.Value = CBool(arr_strSettings(40)) 'objOption.SPEEDMONITOR
                .scrRSMonRefr.Value = arr_strSettings(41) 'objOption.SPEEDMON_REFRESHRATE
                .chkFavAnn.Value = CBool(arr_strSettings(42)) 'objOption.ANNOUNCE_FAV
                .chkAutoSaveRace.Value = CBool(arr_strSettings(43)) 'objOption.AUTO_SAVE
                .optRSMon1.Value = CBool(arr_strSettings(44)) 'objOption.RSMON_SPEED
                .optRSMon2.Value = CBool(arr_strSettings(45)) 'objOption.RSMON_DISTANCE
                .opt_grid1.Value = CBool(arr_strSettings(46)) 'objOption.STARTING_GRID_IN
                .opt_grid2.Value = CBool(arr_strSettings(47)) 'objOption.STARTING_GRID_BEHIND
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

'Save the race log to a text file on the hard disk
Public Sub SaveRaceForReplay(autosave As Boolean)

    On Error GoTo ERRORHANDLING 'In case an error occurs

    Dim varSave As Variant 'Folder and file for saving
    Dim strFileName As String 'File name
    Dim intFileNr As Integer 'Output channel
    Dim strTitle As String 'Pop-up title
    Dim i As Integer, j As Long
    
        strFileName = g_arr_varReplay_RaceData(4, 1) & "-" & g_arr_varReplay_RaceData(5, 1) & "_" _
            & "ID-" & g_arr_varReplay_RaceData(3, 1) & "_" & g_arr_varReplay_RaceData(2, 1)(3) & "-" _
            & g_arr_varReplay_RaceData(2, 1)(2) & "-" & g_arr_varReplay_RaceData(2, 1)(1) & "_" _
            & g_arr_varReplay_RaceData(2, 1)(4) & "h" & g_arr_varReplay_RaceData(2, 1)(5) & "min"
            
        strTitle = GetText(g_arr_Text, "BTN023") & ": " & g_arr_varReplay_RaceData(4, 1) & " (" & g_arr_varReplay_RaceData(5, 1) & ")"

        If Not autosave Then
            varSave = Application.GetSaveAsFilename( _
            InitialFileName:=strFileName, _
            FileFilter:="GaloppSim Races (*.gsrace), *.gsrace", _
            Title:=strTitle)
        Else
            varSave = g_defaultAutoSavePath & Application.PathSeparator & strFileName & ".gsrace"
        End If
        
        If varSave <> False Then
        
            intFileNr = FreeFile 'Assign the next free number
            
            Open varSave For Output As #intFileNr 'Open the output channel
                    'RACE DATA
                    Print #intFileNr, "[GALOPPSIM VERSION]"
                    Print #intFileNr, g_arr_varReplay_RaceData(1, 1) 'GaloppSim Version
                    Print #intFileNr, "[RUN ON]"
                    Print #intFileNr, g_arr_varReplay_RaceData(2, 1)(1) 'Date (run on)
                    Print #intFileNr, g_arr_varReplay_RaceData(2, 1)(2)
                    Print #intFileNr, g_arr_varReplay_RaceData(2, 1)(3)
                    Print #intFileNr, g_arr_varReplay_RaceData(2, 1)(4)
                    Print #intFileNr, g_arr_varReplay_RaceData(2, 1)(5)
                    Print #intFileNr, "[RACE_ID]"
                    Print #intFileNr, g_arr_varReplay_RaceData(3, 1) 'RACE_ID
                    Print #intFileNr, "[PARTICIPANTS]"
                    Print #intFileNr, g_arr_varReplay_RaceData(4, 1) 'PARTICIPANTS
                    Print #intFileNr, "[RACE_NAME]"
                    Print #intFileNr, g_arr_varReplay_RaceData(5, 1) 'RACE_NAME
                    Print #intFileNr, "[RACE_YEAR]"
                    Print #intFileNr, g_arr_varReplay_RaceData(6, 1) 'RACE_YEAR
                    Print #intFileNr, "[TRACK_LOCATION]"
                    Print #intFileNr, g_arr_varReplay_RaceData(7, 1) 'TRACK_LOCATION
                    Print #intFileNr, "[COUNTRY_CODE]"
                    Print #intFileNr, g_arr_varReplay_RaceData(8, 1) 'COUNTRY_CODE
                    If g_arr_varReplay_RaceData(8, 1) = "MOON" Then
                        Print #intFileNr, "[SPACE_PLANET]"
                        Print #intFileNr, g_arr_varReplay_RaceData(24, 1)
                    End If
                    Print #intFileNr, "[TRACK_NAME]"
                    Print #intFileNr, g_arr_varReplay_RaceData(9, 1) 'TRACK_NAME
                    Print #intFileNr, "[TRACK_COLOUR]"
                    Print #intFileNr, g_arr_varReplay_RaceData(10, 1) 'TRACK_COLOUR
                    Print #intFileNr, "[TRACK_SURFACE]"
                    Print #intFileNr, g_arr_varReplay_RaceData(11, 1) 'TRACK_SURFACE
                    Print #intFileNr, "[RACE_TYPE]"
                    Print #intFileNr, g_arr_varReplay_RaceData(12, 1) 'RACE_TYPE
                    Print #intFileNr, "[METRES]"
                    Print #intFileNr, g_arr_varReplay_RaceData(13, 1) 'METRES
                    Print #intFileNr, "[STARTING_GATE]"
                    Print #intFileNr, g_arr_varReplay_RaceData(14, 1) 'STARTING_GATE
                    Print #intFileNr, "[NUMBER_ENROLLED]"
                    Print #intFileNr, g_arr_varReplay_RaceData(15, 1) 'NUMBER_ENROLLED
                    Print #intFileNr, "[NUMBER_STARTING]"
                    Print #intFileNr, g_arr_varReplay_RaceData(16, 1) 'NUMBER_STARTING
                    Print #intFileNr, "[LANES_FIX_OR_RANDOM]"
                    Print #intFileNr, g_arr_varReplay_RaceData(17, 1) 'LANES_FIX_OR_RANDOM
                    Print #intFileNr, "[ADVERTISING]"
                    Print #intFileNr, g_arr_varReplay_RaceData(18, 1) 'ADVERTISING
                    Print #intFileNr, "[SPECIAL]"
                    Print #intFileNr, g_arr_varReplay_RaceData(19, 1) 'SPECIAL
                    Print #intFileNr, "[SPECTATORS]"
                    Print #intFileNr, g_arr_varReplay_RaceData(23, 1) 'SPECTATORS
                    
                    If g_arr_varReplay_RaceData(18, 1) = "Y" Then 'If ADVERTISING
                        Print #intFileNr, "[ADVERTISEMENT LENGTH]"
                        Print #intFileNr, UBound(g_arr_varReplay_RaceData(20, 1)) 'Array length
                        For i = 1 To UBound(g_arr_varReplay_RaceData(20, 1))
                            Print #intFileNr, g_arr_varReplay_RaceData(20, 1)(i) 'Save ADVERTISEMENT
                        Next i
                    Else
                        Print #intFileNr, "[NO ADVERTISEMENT]"
                    End If
                    
                    If g_arr_varReplay_RaceData(21, 1) = "Y" Then 'If TRIBUNES
                        Print #intFileNr, "[TRACK GRAPHICS LENGTH]"
                        Print #intFileNr, UBound(g_arr_varReplay_RaceData(22, 1)) 'Array length
                        For i = 1 To UBound(g_arr_varReplay_RaceData(22, 1))
                            Print #intFileNr, g_arr_varReplay_RaceData(22, 1)(i) 'Save TRACK GRAPHICS
                        Next i
                    Else
                        Print #intFileNr, "[NO TRACK GRAPHICS]"
                    End If

                    'HORSE DATA
                    Print #intFileNr, "[RACE LOOPS]"
                    Print #intFileNr, UBound(g_arr_varReplay_HorseData(1, 14)) 'Race loops

                    For i = 1 To g_arr_varReplay_RaceData(15, 1) 'NUMBER_ENROLLED
                        Print #intFileNr, "[HORSE]"
                        Print #intFileNr, g_arr_varReplay_HorseData(i, 1) 'Name of the horse
                        Print #intFileNr, g_arr_varReplay_HorseData(i, 0) 'Status before the race
                        Print #intFileNr, g_arr_varReplay_HorseData(i, 20) 'Status after the race

                        If Not IsArray(g_arr_varReplay_HorseData(i, 2)) Then 'Horse colour
                            Print #intFileNr, "MONOCHROME"
                            Print #intFileNr, g_arr_varReplay_HorseData(i, 2)
                        Else
                            Print #intFileNr, "MULTICOLOUR"
                            For j = 0 To 7
                                Print #intFileNr, g_arr_varReplay_HorseData(i, 2)(j)
                            Next j
                        End If
                        
                        Print #intFileNr, g_arr_varReplay_HorseData(i, 3) - objBasicData.TOP_ROWS
                                    'Row number on which the horse is running (AI/RS harmonized)
                        Print #intFileNr, g_arr_varReplay_HorseData(i, 11) 'Starting number
                        Print #intFileNr, g_arr_varReplay_HorseData(i, 15) 'Starting gate
                        Print #intFileNr, g_arr_varReplay_HorseData(i, 23) 'Picture of the winner
                        Print #intFileNr, g_arr_varReplay_HorseData(i, 24) 'For SPECIAL purposes
                        Print #intFileNr, g_arr_varReplay_HorseData(i, 25) 'Tactics
                        For j = 1 To UBound(g_arr_varReplay_HorseData(i, 14)) 'Race speed log
                            Print #intFileNr, g_arr_varReplay_HorseData(i, 14)(j)
                        Next j
                    Next i
            Close #intFileNr 'Close the output channel
            
            If Not autosave Then
                'Show a pop-up: "Race data successfully saved."
                Call ShowMessagePopup(GetText(g_arr_Text, "BTN023"), _
                    GetText(g_arr_Text, "SUCCESS002") & vbNewLine & vbNewLine & varSave, _
                    enumButton.OK, vbModal)
            End If
        End If
        
'Save GEMA data file
        If Not autosave Then
            varSave = left(varSave, Len(varSave) - 7) & ".gsgema"
        Else
            varSave = g_defaultAutoSavePath & Application.PathSeparator & strFileName & ".gsgema"
        End If
        
        If varSave <> False Then
        
            intFileNr = FreeFile 'Assign the next free number
            
            Open varSave For Output As #intFileNr 'Open the output channel
                    'GEMA DATA
                    Print #intFileNr, "[GEMA ROWS]"
                    Print #intFileNr, UBound(g_arr_GEMA()) 'GEMA rows
                    Print #intFileNr, "[GEMA LENGTH]"
                    Print #intFileNr, UBound(g_arr_GEMA(), 2) 'GEMA columns

                    For i = 1 To UBound(g_arr_GEMA()) 'GEMA rows
                        For j = 1 To UBound(g_arr_GEMA(), 2) 'GEMA columns
                            Print #intFileNr, g_arr_GEMA(i, j)
                        Next j
                    Next i

            Close #intFileNr 'Close the output channel
            
            
            If Not autosave Then
                'Show a pop-up: "GEMA data successfully saved."
                Call ShowMessagePopup(GetText(g_arr_Text, "BTN023"), _
                    GetText(g_arr_Text, "SUCCESS004") & vbNewLine & vbNewLine & varSave, _
                    enumButton.OK, vbModal)
            End If
        End If
        
'Save tunnel data file
        If objRace.SPECIAL = "TUNNEL" Then
            If Not autosave Then
                varSave = left(varSave, Len(varSave) - 7) & ".gstunn"
            Else
                varSave = g_defaultAutoSavePath & Application.PathSeparator & strFileName & ".gstunn"
            End If
            
            If varSave <> False Then
            
                intFileNr = FreeFile 'Assign the next free number
                
                Open varSave For Output As #intFileNr 'Open the output channel
                        'ROAD CROSSING AND POLICE CAR
                        Print #intFileNr, "[ROADCROSSING]"
                        Print #intFileNr, objOption.ROADCROSSING
                        Print #intFileNr, "[POLICECAR]"
                        Print #intFileNr, objOption.POLICECAR
                        'TUNNEL DATA
                        Print #intFileNr, "[TUNNEL DATA DIMENSION 1]"
                        Print #intFileNr, UBound(g_arr_varReplay_TunnelData())
                        Print #intFileNr, "[TUNNEL DATA DIMENSION 2]"
                        Print #intFileNr, UBound(g_arr_varReplay_TunnelData(), 2)
    
                        For i = 1 To UBound(g_arr_varReplay_TunnelData())
                            For j = 1 To UBound(g_arr_varReplay_TunnelData(), 2)
                                Print #intFileNr, g_arr_varReplay_TunnelData(i, j)
                            Next j
                        Next i
    
                Close #intFileNr 'Close the output channel
                
                
                If Not autosave Then
                    'Show a pop-up: "Tunnel data successfully saved."
                    Call ShowMessagePopup(GetText(g_arr_Text, "BTN023"), _
                        GetText(g_arr_Text, "SUCCESS005") & vbNewLine & vbNewLine & varSave, _
                        enumButton.OK, vbModal)
                End If
            End If
        End If

    Exit Sub
ERRORHANDLING:
    'Show a pop-up: "Error while saving."
    Call ShowMessagePopup(UCase(GetText(g_arr_Text, "BTN023")), _
        GetText(g_arr_Text, "ERROR009"), _
        enumButton.OK, vbModal)
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "SaveRaceForReplay()")
    Close #intFileNr 'Close the output channel
End Sub

'Load a race log from a text file on the hard disk
Public Sub LoadRaceForReplay()

    On Error GoTo ERRORHANDLING 'In case an error occurs

    Dim varLoad As Variant 'Folder and file for loading
    Dim intFileNr As Integer 'Input channel
    Dim strTitle As String 'Pop-up title
    Dim strInput As String
    Dim i As Integer, j As Long
    Dim tempHorses As Integer, tempLoops As Long
    Dim tempArray() As Variant
    Dim tempRacelog() As Double
    Dim gemaRows As Integer
    Dim gemaLength As Integer
    Dim tunnelDim1 As Integer
    Dim tunnelDim2 As Integer
    
    ReDim g_arr_varReplay_RaceData(1 To 24, 0 To 1)
    g_arr_varReplay_RaceData(1, 0) = "GaloppSim Version"
    g_arr_varReplay_RaceData(2, 0) = "Date (run on)"
    g_arr_varReplay_RaceData(3, 0) = "RACE_ID"
    g_arr_varReplay_RaceData(4, 0) = "PARTICIPANTS"
    g_arr_varReplay_RaceData(5, 0) = "RACE_NAME"
    g_arr_varReplay_RaceData(6, 0) = "RACE_YEAR"
    g_arr_varReplay_RaceData(7, 0) = "TRACK_LOCATION"
    g_arr_varReplay_RaceData(8, 0) = "COUNTRY_CODE"
    g_arr_varReplay_RaceData(24, 0) = "SPACE_PLANET"
    g_arr_varReplay_RaceData(9, 0) = "TRACK_NAME"
    g_arr_varReplay_RaceData(10, 0) = "TRACK_COLOUR"
    g_arr_varReplay_RaceData(11, 0) = "TRACK_SURFACE"
    g_arr_varReplay_RaceData(12, 0) = "RACE_TYPE"
    g_arr_varReplay_RaceData(13, 0) = "METRES"
    g_arr_varReplay_RaceData(14, 0) = "STARTING_GATE"
    g_arr_varReplay_RaceData(15, 0) = "NUMBER_ENROLLED"
    g_arr_varReplay_RaceData(16, 0) = "NUMBER_STARTING"
    g_arr_varReplay_RaceData(17, 0) = "LANES_FIX_OR_RANDOM"
    g_arr_varReplay_RaceData(18, 0) = "ADVERTISING"
    g_arr_varReplay_RaceData(19, 0) = "SPECIAL"
    g_arr_varReplay_RaceData(20, 0) = "ADVERTISEMENT"
    g_arr_varReplay_RaceData(21, 0) = "TRIBUNES"
    g_arr_varReplay_RaceData(22, 0) = "TRACK GRAPHICS"
    g_arr_varReplay_RaceData(23, 0) = "SPECTATORS"
    
    strTitle = GetText(g_arr_Text, "BTN024")

    varLoad = Application.GetOpenFilename( _
    FileFilter:="GaloppSim Races (*.gsrace), *.gsrace", Title:=strTitle)
    
    If varLoad = False Then Exit Sub
    
    If Right(varLoad, 7) = ".gsrace" Then 'Check whether the file extension is "gsrace"
    
        intFileNr = FreeFile 'Assign the next free number

        'Load the data
        Open varLoad For Input As #intFileNr 'Open the input channel
            'RACE DATA
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(1, 1) = strInput 'GaloppSim version
            ReDim tempArray(1 To 5)
            Line Input #intFileNr, strInput '(description)
            For i = 1 To 5
                Line Input #intFileNr, strInput
                tempArray(i) = CInt(strInput)
            Next i
                g_arr_varReplay_RaceData(2, 1) = tempArray 'Date (run on)
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(3, 1) = strInput 'RACE_ID
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(4, 1) = strInput 'PARTICIPANTS
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(5, 1) = strInput 'RACE_NAME
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(6, 1) = strInput 'RACE_YEAR
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(7, 1) = strInput 'TRACK_LOCATION
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(8, 1) = strInput 'COUNTRY_CODE
            If g_arr_varReplay_RaceData(8, 1) = "MOON" Then
                Line Input #intFileNr, strInput '(description)
                Line Input #intFileNr, strInput
                    g_arr_varReplay_RaceData(24, 1) = strInput 'SPACE_PLANET
            End If
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(9, 1) = strInput 'TRACK_NAME
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(10, 1) = CLng(strInput) 'TRACK_COLOUR
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(11, 1) = strInput 'TRACK_SURFACE
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(12, 1) = strInput 'RACE_TYPE
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(13, 1) = CLng(strInput) 'METRES
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(14, 1) = strInput 'STARTING_GATE
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(15, 1) = CInt(strInput) 'NUMBER_ENROLLED
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(16, 1) = CInt(strInput) 'NUMBER_STARTING
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(17, 1) = strInput 'LANES_FIX_OR_RANDOM
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(18, 1) = strInput 'ADVERTISING
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(19, 1) = strInput 'SPECIAL
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                g_arr_varReplay_RaceData(23, 1) = strInput 'SPECTATORS

            Line Input #intFileNr, strInput '(description)
            If g_arr_varReplay_RaceData(18, 1) = "Y" Then 'If ADVERTISING
                Line Input #intFileNr, strInput
                j = CStr(strInput) 'Array length
                ReDim tempArray(1 To j)
                For i = 1 To j
                    Line Input #intFileNr, strInput
                        tempArray(i) = strInput 'ADVERTISEMENT
                Next i
                g_arr_varReplay_RaceData(20, 1) = tempArray
            End If

            Line Input #intFileNr, strInput '(description)
                If strInput = "[TRACK GRAPHICS LENGTH]" Then 'TRIBUNES
                    Line Input #intFileNr, strInput
                        j = CStr(strInput) 'Array length
                    ReDim tempArray(1 To j)
                    For i = 1 To j
                        Line Input #intFileNr, strInput
                            tempArray(i) = strInput 'TRACK GRAPHICS
                    Next i
                    g_arr_varReplay_RaceData(22, 1) = tempArray
                    objOption.TRIBUNES = True
                Else
                    objOption.TRIBUNES = False
                End If

            'HORSE DATA
            Line Input #intFileNr, strInput '(description)
            Line Input #intFileNr, strInput
                tempLoops = strInput 'Race loops
                
            tempHorses = g_arr_varReplay_RaceData(15, 1) 'NUMBER_ENROLLED

            ReDim g_arr_varReplay_HorseData(1 To tempHorses, 0 To 31)
            For i = 1 To tempHorses
                Line Input #intFileNr, strInput '(description)
                Line Input #intFileNr, strInput
                    g_arr_varReplay_HorseData(i, 1) = strInput 'Name of the horse
                Line Input #intFileNr, strInput
                    g_arr_varReplay_HorseData(i, 0) = strInput 'Status before the race
                Line Input #intFileNr, strInput
                    g_arr_varReplay_HorseData(i, 20) = strInput 'Status after the race
                    
                Line Input #intFileNr, strInput
                    If strInput = "MONOCHROME" Then
                        Line Input #intFileNr, strInput
                            g_arr_varReplay_HorseData(i, 2) = CLng(strInput)
                    Else
                        ReDim tempArray(0 To 7)
                        For j = 0 To 7
                            Line Input #intFileNr, strInput
                            tempArray(j) = CLng(strInput)
                        Next j
                        g_arr_varReplay_HorseData(i, 2) = tempArray
                    End If
                    
                Line Input #intFileNr, strInput
                    g_arr_varReplay_HorseData(i, 3) = CInt(strInput) + objBasicData.TOP_ROWS
                            'Row number on which the horse is running (AI/RS harmonized)
                Line Input #intFileNr, strInput
                    g_arr_varReplay_HorseData(i, 11) = CDbl(strInput) 'Starting number
                Line Input #intFileNr, strInput
                    g_arr_varReplay_HorseData(i, 15) = CInt(strInput) 'Starting gate
                Line Input #intFileNr, strInput
                    g_arr_varReplay_HorseData(i, 23) = strInput 'Picture of the winner
                Line Input #intFileNr, strInput
                    g_arr_varReplay_HorseData(i, 24) = strInput 'For SPECIAL purposes
                Line Input #intFileNr, strInput
                    g_arr_varReplay_HorseData(i, 25) = strInput 'Tactics
                    
                ReDim tempRacelog(1 To tempLoops)
                For j = 1 To tempLoops
                    Line Input #intFileNr, strInput
                        tempRacelog(j) = CDbl(strInput)
                Next j
                g_arr_varReplay_HorseData(i, 14) = tempRacelog
            Next i
        Close #intFileNr 'Close the input channel

    Else 'Wrong file format
        Call ShowMessagePopup(GetText(g_arr_Text, "ERROR010"), _
            GetText(g_arr_Text, "ERROR011"), _
            enumButton.OK, vbModal)
        Exit Sub
    End If
    
'Load GEMA data file (if the file exists)
    #If Debugging Then
        Debug.Print "Search for a GEMA file (.gsgema)..."
    #End If

    varLoad = left(varLoad, Len(varLoad) - 7) & ".gsgema"
    
    If Dir(varLoad) <> "" Then
    #If Debugging Then
        Debug.Print "GEMA file found. Load GEMA data."
    #End If

        intFileNr = FreeFile 'Assign the next free number

        'Load the data
        Open varLoad For Input As #intFileNr 'Open the input channel
            Line Input #intFileNr, strInput 'description [GEMA ROWS]
            Line Input #intFileNr, strInput
                gemaRows = strInput
            Line Input #intFileNr, strInput 'description [GEMA LENGTH]
            Line Input #intFileNr, strInput
                gemaLength = strInput
            
            ReDim g_arr_GEMA(1 To gemaRows, 1 To gemaLength)
            
            For i = 1 To gemaRows
                For j = 1 To gemaLength
                    Line Input #intFileNr, strInput
                    g_arr_GEMA(i, j) = strInput
                Next j
            Next i
        
        Close #intFileNr 'Close the input channel
    
    Else
    #If Debugging Then
        Debug.Print "No GEMA file found."
    #End If
    End If
    
'Load tunnel data file (if the file exists)
    #If Debugging Then
        Debug.Print "Search for a TUNNEL file (.gstunn)..."
    #End If
    varLoad = left(varLoad, Len(varLoad) - 7) & ".gstunn"
    
    If Dir(varLoad) <> "" Then
    #If Debugging Then
        Debug.Print "TUNNEL file found. Load tunnel data."
    #End If

        intFileNr = FreeFile 'Assign the next free number

        'Load the data
        Open varLoad For Input As #intFileNr 'Open the input channel
            Line Input #intFileNr, strInput 'description [ROADCROSSING]
            Line Input #intFileNr, strInput
                objOption.ROADCROSSING = CBool(strInput)
            Line Input #intFileNr, strInput 'description [POLICECAR]
            Line Input #intFileNr, strInput
                objOption.POLICECAR = CBool(strInput)
            Line Input #intFileNr, strInput 'description [TUNNEL DATA DIMENSION 1]
            Line Input #intFileNr, strInput
                tunnelDim1 = strInput
            Line Input #intFileNr, strInput 'description [TUNNEL DATA DIMENSION 2]
            Line Input #intFileNr, strInput
                tunnelDim2 = strInput

            ReDim g_arr_varReplay_TunnelData(1 To tunnelDim1, 0 To tunnelDim2)

            For i = 1 To tunnelDim1
                For j = 1 To tunnelDim2
                    Line Input #intFileNr, strInput
                    g_arr_varReplay_TunnelData(i, j) = strInput
                Next j
            Next i
        
        Close #intFileNr 'Close the input channel
    
    Else
    #If Debugging Then
        Debug.Print "No TUNNEL file found."
    #End If
    End If
    
    
    objRace.LOADED = True
    
    Call ShowMessagePopup("", GetText(g_arr_Text, "SUCCESS003") & ". " _
        & GetText(g_arr_Text, "BTN025") & "?", enumButton.YesNo, vbModal)
    If g_enumButton = enumButton.yes Then Call RaceReplay

    If g_strPlayMode = "RS" Then
        g_wksRace.OLEObjects("replay").Object.Enabled = True
    Else
        g_RibbonGaloppSim.Invalidate 'Refresh the menu buttons
    End If
    
    Exit Sub
ERRORHANDLING:
    'Show a pop-up: "Error while loading."
    Call ShowMessagePopup(UCase(GetText(g_arr_Text, "BTN024")), _
        GetText(g_arr_Text, "ERROR010"), _
        enumButton.OK, vbModal)
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "LoadRaceForReplay()")
    Close #intFileNr 'Close the input channel
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
