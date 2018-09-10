Attribute VB_Name = "basInputOutput"
Option Explicit

Public Sub SaveRaceOptions()

    Dim strFile As String 'file name and path
    Dim intFileNr As Integer 'output channel
    
    On Error GoTo ERRORHANDLING
    
        strFile = g_defaultPath & g_c_defaultRaceOptionsFile & g_c_defaultFileType
        intFileNr = FreeFile 'assign next free number
        
        Open strFile For Output As #intFileNr 'open the output channel
            With frmOptionsRace
                'Write settings into a text file
                Print #intFileNr, CInt(.optRS_ro01.Value) 'g_blnTactics0
                Print #intFileNr, CInt(.optRS_ro02.Value) 'g_blnTactics3
                Print #intFileNr, CInt(.optRS_ro03.Value) 'g_blnTactics6
                Print #intFileNr, CInt(.chkRS_ro12a.Value) 'g_blnSlipstream
                Print #intFileNr, CInt(.chkRS_ro12b.Value) 'g_blnSlipstreamDouble
                Print #intFileNr, CInt(.chkRS_ro12c.Value) 'g_blnSlipstreamShow
                Print #intFileNr, CInt(.chkRS_ro01a.Value) 'g_blnFocusedRun
                Print #intFileNr, CInt(.chkRS_ro01b.Value) 'g_blnHighlightFoc
                Print #intFileNr, CInt(.chkRS_ro11a.Value) 'g_blnBettingMode
                Print #intFileNr, CInt(.chkRS_ro11b.Value) 'g_blnBettingAnalysis
                Print #intFileNr, CInt(.chkRS_ro10.Value) 'g_blnRaceInformation
                Print #intFileNr, CInt(.chkRS_ro13.Value) 'g_blnRaceInfoLeader
                Print #intFileNr, CInt(.chkRS_ro14.Value) 'g_blnRaceInfoProgressBar
                Print #intFileNr, .lblRS_ro07.BackColor 'g_lngRaceInfoBackColour
                Print #intFileNr, .lblRS_ro07.ForeColor 'g_lngRaceInfoForeColour
                Print #intFileNr, CInt(.chkRS_ro02.Value) 'g_blnHoofprints
                Print #intFileNr, .scrRS_ro01.Value 'g_byteTrackZoom
                Print #intFileNr, .scrRS_ro02.Value 'g_intTrackMetres
                Print #intFileNr, CInt(.chkRS_ro04.Value) 'g_blnHorseNamesLeft
                Print #intFileNr, CInt(.chkRS_ro05.Value) 'g_blnHorseColoursLeft
                Print #intFileNr, CInt(.chkRS_ro06.Value) 'g_blnHighlightFav
                Print #intFileNr, CInt(.chkRS_ro07.Value) 'g_blnHorseNamesFinish
                Print #intFileNr, CInt(.chkRS_ro08.Value) 'g_blnRankingColours
                Print #intFileNr, CInt(.chkRS_ro09.Value) 'g_blnRankingDelay
                Print #intFileNr, CInt(.opt_ro01.Value) 'g_blnRaceInfoPopup
                Print #intFileNr, CInt(.opt_ro02.Value) 'g_blnRaceInfoWorksheet
                Print #intFileNr, .scrRS_ro03.Value 'g_intGMPspeedfactor
                Print #intFileNr, CInt(.chkRS_ro15.Value) 'g_blnIncidentRefuse
                Print #intFileNr, CInt(.chkRS_ro16.Value) 'g_blnSpeech
            End With
        Close #intFileNr 'close output channel

        'Show pop-up
            'Set the button mode
            g_strMsgButtons = "OK"
            'Assign the text for the pop-up
            g_strMsgCaption = GetTxt(g_arrTxt, "BTN001") & " - " & GetTxt(g_arrTxt, "RACEOPT037")
            g_strMsgText = GetTxt(g_arrTxt, "SUCCESS001") & vbNewLine & vbNewLine & strFile
            'Display the pop-up
            frmMsg_MultiPurpose.Show (vbModal) 'modal
                
    Exit Sub
    
ERRORHANDLING:
    'Show pop-up
        'Set the button mode
        g_strMsgButtons = "OK"
        'Assign the text for the pop-up
        g_strMsgCaption = GetTxt(g_arrTxt, "BTN001") & " - " & GetTxt(g_arrTxt, "RACEOPT037")
        g_strMsgText = GetTxt(g_arrTxt, "ERROR005")
        'Display the pop-up
        frmMsg_MultiPurpose.Show (vbModal) 'modal
End Sub

Public Sub LoadRaceOptions()

    Dim strFile As String 'file name and path
    Dim intFileNr As Integer 'output channel
    Dim arr_strSettings(1 To 29) As String 'string array for the values
    Dim i As Integer
    
    On Error Resume Next 'ignore errors
    
    strFile = g_defaultPath & g_c_defaultRaceOptionsFile & g_c_defaultFileType
    
    If Dir(strFile) <> "" Then
    
        intFileNr = FreeFile 'assign next free number
        
        'Load settings and write them into an array
        Open strFile For Input As #intFileNr 'open the input channel
            For i = 1 To UBound(arr_strSettings)
                Line Input #intFileNr, arr_strSettings(i)
            Next
        Close #intFileNr 'close the input channel
        
        'Write the settings into the race options UserForm
            With frmOptionsRace
                .optRS_ro01.Value = CBool(arr_strSettings(1)) 'g_blnTactics0
                .optRS_ro02.Value = CBool(arr_strSettings(2)) 'g_blnTactics3
                .optRS_ro03.Value = CBool(arr_strSettings(3)) 'g_blnTactics6
                .chkRS_ro12a.Value = CBool(arr_strSettings(4)) 'g_blnSlipstream
                .chkRS_ro12b.Value = CBool(arr_strSettings(5)) 'g_blnSlipstreamDouble
                .chkRS_ro12c.Value = CBool(arr_strSettings(6)) 'g_blnSlipstreamShow
                .chkRS_ro01a.Value = CBool(arr_strSettings(7)) 'g_blnFocusedRun
                .chkRS_ro01b.Value = CBool(arr_strSettings(8)) 'g_blnHighlightFoc
                .chkRS_ro11a.Value = CBool(arr_strSettings(9)) 'g_blnBettingMode
                .chkRS_ro11b.Value = CBool(arr_strSettings(10)) 'g_blnBettingAnalysis
                .chkRS_ro10.Value = CBool(arr_strSettings(11)) 'g_blnRaceInformation
                .chkRS_ro13.Value = CBool(arr_strSettings(12)) 'g_blnRaceInfoLeader
                .chkRS_ro14.Value = CBool(arr_strSettings(13)) 'g_blnRaceInfoProgressBar
                .lblRS_ro07.BackColor = arr_strSettings(14) 'g_lngRaceInfoBackColour
                .lblRS_ro07.ForeColor = arr_strSettings(15) 'g_lngRaceInfoForeColour
                .chkRS_ro02.Value = CBool(arr_strSettings(16)) 'g_blnHoofprints
                .scrRS_ro01.Value = arr_strSettings(17) 'g_byteTrackZoom
                .scrRS_ro02.Value = arr_strSettings(18) 'g_intTrackMetres
                .chkRS_ro04.Value = CBool(arr_strSettings(19)) 'g_blnHorseNamesLeft
                .chkRS_ro05.Value = CBool(arr_strSettings(20)) 'g_blnHorseColoursLeft
                .chkRS_ro06.Value = CBool(arr_strSettings(21)) 'g_blnHighlightFav
                .chkRS_ro07.Value = CBool(arr_strSettings(22)) 'g_blnHorseNamesFinish
                .chkRS_ro08.Value = CBool(arr_strSettings(23)) 'g_blnRankingColours
                .chkRS_ro09.Value = CBool(arr_strSettings(24)) 'g_blnRankingDelay
                .opt_ro01.Value = CBool(arr_strSettings(25)) 'g_blnRaceInfoPopup
                .opt_ro02.Value = CBool(arr_strSettings(26)) 'g_blnRaceInfoWorksheet
                .scrRS_ro03.Value = arr_strSettings(27) 'g_intGMPspeedfactor
                .chkRS_ro15.Value = CBool(arr_strSettings(28)) 'g_blnIncidentRefuse
                .chkRS_ro16.Value = CBool(arr_strSettings(29)) 'g_blnSpeech
            End With
    Else
        'Show pop-up
            'Set the button mode
            g_strMsgButtons = "OK"
            'Assign the text for the pop-up
            g_strMsgCaption = GetTxt(g_arrTxt, "BTN001") & " - " & GetTxt(g_arrTxt, "RACEOPT038")
            g_strMsgText = GetTxt(g_arrTxt, "ERROR006")
            'Display the pop-up
            frmMsg_MultiPurpose.Show (vbModal) 'modal
    End If
    
    On Error GoTo 0
End Sub

Public Sub SpeechOut(words As String)
    Application.speech.Speak (words), SpeakAsync:=True
End Sub
