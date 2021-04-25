VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBetSlip 
   Caption         =   "[Name]"
   ClientHeight    =   8952
   ClientLeft      =   -1548
   ClientTop       =   -6876
   ClientWidth     =   20124
   OleObjectBlob   =   "frmBetSlip.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmBetSlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up with a betting slip
'   UserForm frmBetSlip

Dim m_arrintBetSlip(1 To 4, 1 To 24) As Integer 'Array for the entire bet slip
Dim m_arrintBet() As Integer 'Array for the ticked Checkboxes
Dim m_dblStake As Double 'Stake
Dim m_dblOdds As Double 'Odds
Dim m_lngType As Long 'Type of the bet (ID)
Dim m_strType_Text As String 'Type of the bet (text)
Dim m_blnOK As Boolean 'True when bet slip is valid

Private Sub UserForm_Initialize() 'Set default values
    'Section: Race data
    lbl110.caption = GetText(g_arr_Text, "BET005") & ":"
    lbl112.caption = GetText(g_arr_Text, "BET003")
    lbl113.caption = GetText(g_arr_Text, "BET004")
    'Section: Checkboxes for place 1-4
    lbl115.caption = g_arr_Grammar(3)
    'Section: Type of bet
    fraBettingType.caption = GetText(g_arr_Text, "BET007")
    opt117.caption = GetText(g_arr_Text, "BET008")
    opt117.ControlTipText = GetText(g_arr_Text, "BET051")
    opt118.caption = GetText(g_arr_Text, "BET009")
    opt118.ControlTipText = GetText(g_arr_Text, "BET052")
    opt120.caption = GetText(g_arr_Text, "BET010")
    opt120.ControlTipText = GetText(g_arr_Text, "BET053")
    opt121.caption = GetText(g_arr_Text, "BET011")
    opt121.ControlTipText = GetText(g_arr_Text, "BET054")
    opt122.caption = GetText(g_arr_Text, "BET012")
    opt122.ControlTipText = GetText(g_arr_Text, "BET055")
    opt123.caption = GetText(g_arr_Text, "BET013")
    opt123.ControlTipText = GetText(g_arr_Text, "BET056")
    'Section: Stake
    fraStake.caption = GetText(g_arr_Text, "BET014")
    opt125.caption = GetText(g_arr_Text, "BET015")
    opt126.caption = GetText(g_arr_Text, "BET016")
    opt127.caption = GetText(g_arr_Text, "BET017")
    opt128.caption = GetText(g_arr_Text, "BET018")
    opt129.caption = GetText(g_arr_Text, "BET019")
    opt130.caption = GetText(g_arr_Text, "BET020")
    opt131.caption = GetText(g_arr_Text, "BET021")
    opt132.caption = GetText(g_arr_Text, "BET022")
    opt133.caption = GetText(g_arr_Text, "BET023")
    opt134.caption = GetText(g_arr_Text, "BET024")
    opt135.caption = GetText(g_arr_Text, "BET025")
    'Buttons
    cmd140.caption = GetText(g_arr_Text, "START008") '"Show horse numbers and odds"
    cmd141.caption = GetText(g_arr_Text, "BET026") '"Place bet"
    cmd142.caption = GetText(g_arr_Text, "BET027") '"Discard betting slip"
    
    'Type of bets
    Call opt117_Click   'Default: Win
    If objRace.NUMBER_STARTING >= 4 Then opt118.Enabled = True _
        Else opt118.Enabled = False 'Show
    If objRace.NUMBER_STARTING >= 10 Then opt121.Enabled = True _
        Else opt121.Enabled = False '2 sur 4
    opt120.Enabled = True 'Exacta
    opt122.Enabled = True 'Trifecta
    opt123.Enabled = True 'Superfecta
    opt127.Value = True 'Default stake: 2 EUR
    
    'Display the UserForm in the center of the window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

'Button SHOW HORSE NUMBERS AND ODDS
Private Sub cmd140_Click()
    Call ShowSpeed(True) 'True --> show odds
End Sub

'Button PLACE BET
Private Sub cmd141_Click()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Dim BetSlipN As clsBetSlip
    
    m_blnOK = True 'Initialize the variable
    
    Call ReadBetSlip
    Call ValidateBetSlip
    
    If m_blnOK = True Then 'If bet slip is valid
        Set BetSlipN = New clsBetSlip 'Create an instance of this class
        
        With BetSlipN 'Write the values to the properties of the bet slip object
            .BET_ID = CStr(g_colBetSlips.count + 1001) & objRace.RACE_ID 'Compile a unique bet slip ID
            .GAMBLERNAME = Me.caption
            .STAKE = m_dblStake
            .ODDS = m_dblOdds
            .BET_TYPE = m_lngType
            .BET_TYPE_TEXT = m_strType_Text
            .BET = m_arrintBet()
        End With
        
        g_colBetSlips.Add BetSlipN 'Add the bet slip to the Collection
        frmStart.lstBetSlips.AddItem BetSlipN.GAMBLERNAME & " - " & Format(BetSlipN.STAKE, "0.00") & " " & GetText(g_arr_Text, "BET035") _
                                        & " (" & BetSlipN.BET_TYPE_TEXT & ") - [" & GetText(g_arr_Text, "BET001") & " #" & BetSlipN.BET_ID & "]" 'Add bet slip to the ListBox
                                        
        Unload Me 'Close the UserForm with the bet slip
        
        objOption.BET_PLACED = True 'Note that at least one bet slip has been submitted
        Call GetNumberBetSlips 'Refresh the number of bet slips
        frmStart.lstBetSlips.Visible = True 'Show the area with the bet slips
        frmStart.lblBet02.Visible = True 'Show the label with the number of betting slips
    End If
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "cmd141_Click()")
    Call basAuxiliary.CodeCrash
End Sub

'Read all information filled in by the gambler
Private Sub ReadBetSlip()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Dim ctr As control
    Dim i As Integer
    
    'Get the ticked Checkboxes of each row
    i = 1
    For Each ctr In fraPlace1.Controls 'Top row ("I")
        m_arrintBetSlip(1, i) = Abs(ctr.Value * 1)
        i = i + 1
        'Comment in the following lines for understanding the possible boolean values of a Checkbox
'            Debug.Print vbNewLine & "Checkbox number " & i - 1 & " value: " & ctr.Value 'True if ticked, False if not
'            Debug.Print "Checkbox value multiplied by 1: " & ctr.Value * 1 '-1 if true, 0 if false
'            Debug.Print "Converted to an absolute value: " & Abs(ctr.Value * 1) '1 if true, 0 if false
    Next ctr

    i = 1
    For Each ctr In fraPlace2.Controls 'Second row ("II")
        m_arrintBetSlip(2, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr
    
    i = 1
    For Each ctr In fraPlace3.Controls 'Third row ("III")
        m_arrintBetSlip(3, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr
    
        i = 1
    For Each ctr In fraPlace4.Controls 'Fourth row ("IV")
        m_arrintBetSlip(4, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr
    
    'Get the Radio Button with the selected stake
    Select Case True
        Case opt125
            m_dblStake = CDbl(GetText(g_arr_Text, "BET015"))
        Case opt126
            m_dblStake = CDbl(GetText(g_arr_Text, "BET016"))
        Case opt127
            m_dblStake = CDbl(GetText(g_arr_Text, "BET017"))
        Case opt128
            m_dblStake = CDbl(GetText(g_arr_Text, "BET018"))
        Case opt129
            m_dblStake = CDbl(GetText(g_arr_Text, "BET019"))
        Case opt130
            m_dblStake = CDbl(GetText(g_arr_Text, "BET020"))
        Case opt131
            m_dblStake = CDbl(GetText(g_arr_Text, "BET021"))
        Case opt132
            m_dblStake = CDbl(GetText(g_arr_Text, "BET022"))
        Case opt133
            m_dblStake = CDbl(GetText(g_arr_Text, "BET023"))
        Case opt134
            m_dblStake = CDbl(GetText(g_arr_Text, "BET024"))
        Case opt135
            m_dblStake = CDbl(GetText(g_arr_Text, "BET025"))
    End Select
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "ReadBetSlip()")
    Call basAuxiliary.CodeCrash
End Sub

'Check if the entire betting slip is filled in correctly
'depending on the type of bet
Private Sub ValidateBetSlip()
    On Error GoTo ERRORHANDLING 'In case an error occurs

    Select Case True 'Get the selected type of bet
        Case opt117 'Win
            ReDim m_arrintBet(1 To 1)
            m_arrintBet(1) = CheckRow(1, 1)(1) 'Check row I
            #If Debugging Then
                If m_blnOK Then
                    Debug.Print vbNewLine & "Bet on horse number " & m_arrintBet(1) & " - Odds " & FindHorse(m_arrintBet(1)) & ":10"
                Else
                    Debug.Print vbNewLine & "Bet slip is not valid"
                End If
            #End If
            m_dblOdds = FindHorse(m_arrintBet(1)) / 10 'Calculate the payout for 1 EUR stake
            m_lngType = enumBetType.win
            m_strType_Text = GetText(g_arr_Text, "BET008")
            If opt125 Or opt126 Then
                Call ErrorMinStake(2) 'If a stake less than 2 EUR is selected
                opt127.Value = True
            End If

        Case opt118 'Show
            ReDim m_arrintBet(1 To 1)
            m_arrintBet(1) = CheckRow(1, 1)(1) 'Check row I
            m_dblOdds = Round(Application.WorksheetFunction.max( _
                (FindHorse(m_arrintBet(1)) / objRace.NUMBER_STARTING + objRace.NUMBER_STARTING) / 10, _
                1.1), 1) 'Calculate the payout for 1 EUR stake (minimum payout 1.10 EUR)
            m_lngType = enumBetType.show
            m_strType_Text = GetText(g_arr_Text, "BET009")
            If opt125 Or opt126 Then
                Call ErrorMinStake(2)  'If a stake less than 2 EUR is selected
                opt127.Value = True
            End If

        Case opt121 '2 sur 4
            ReDim m_arrintBet(1 To 2)
            m_arrintBet(1) = CheckRow(1, 2)(1) 'Check row I (2 ticks!)
            m_arrintBet(2) = CheckRow(1, 2)(2) 'Check row I (2 ticks!)
            m_dblOdds = 0 'This pay-out has to be calculated after the race!
            m_lngType = enumBetType.x2sur4
            m_strType_Text = GetText(g_arr_Text, "BET011")
            If opt125 Or opt126 Or opt127 Then
                Call ErrorMinStake(3)
                opt128.Value = True
            End If

        Case opt120 'Exacta
            ReDim m_arrintBet(1 To 2)
            m_arrintBet(1) = CheckRow(1, 1)(1) 'Check row I
            m_arrintBet(2) = CheckRow(2, 1)(1) 'Check row II
            m_dblOdds = (FindHorse(m_arrintBet(1)) + FindHorse(m_arrintBet(2))) _
                * 15 / 8 / 10
            m_lngType = enumBetType.exacta
            m_strType_Text = GetText(g_arr_Text, "BET010")
            If opt125 Then
                Call ErrorMinStake(1)
                opt126.Value = True
            End If

        Case opt122 'Trifecta
            ReDim m_arrintBet(1 To 3)
            m_arrintBet(1) = CheckRow(1, 1)(1) 'Check row I
            m_arrintBet(2) = CheckRow(2, 1)(1) 'Check row II
            m_arrintBet(3) = CheckRow(3, 1)(1) 'Check row III
            m_dblOdds = (FindHorse(m_arrintBet(1)) + FindHorse(m_arrintBet(2)) + FindHorse(m_arrintBet(3))) _
                * 25 / 8 / 10
            m_lngType = enumBetType.trifecta
            m_strType_Text = GetText(g_arr_Text, "BET012")

        Case opt123 'Superfecta
            ReDim m_arrintBet(1 To 4)
            m_arrintBet(1) = CheckRow(1, 1)(1) 'Check row I
            m_arrintBet(2) = CheckRow(2, 1)(1) 'Check row II
            m_arrintBet(3) = CheckRow(3, 1)(1) 'Check row III
            m_arrintBet(4) = CheckRow(4, 1)(1) 'Check row IV
            m_dblOdds = (FindHorse(m_arrintBet(1)) + FindHorse(m_arrintBet(2)) + FindHorse(m_arrintBet(3)) _
                + FindHorse(m_arrintBet(4)))
            m_dblOdds = m_dblOdds * 100 / 8 / 10
            m_lngType = enumBetType.superfecta
            m_strType_Text = GetText(g_arr_Text, "BET013")
    End Select
    
    'Check if a horse does not start
    Dim i As Integer, j As Integer
    For i = 1 To UBound(m_arrintBet)
        For j = 1 To UBound(g_arr_varHorses())
            If m_arrintBet(i) = g_arr_varHorses(j, 11) _
                And (g_arr_varHorses(j, 0) = "CANCELLED" Or g_arr_varHorses(j, 0) = "CORONAVIRUSPOSITIVE") Then
                    Call ShowInfoPopup(g_c_tool, _
                        GetText(g_arr_Text, "BET029") & " " & m_arrintBet(i) & " " & GetText(g_arr_Text, "BET058"), True, _
                        vbModal, 12)
                    m_blnOK = False 'Bet slip is not valid
                    Exit For
            End If
        Next
    Next
    
    'Check if the same horse is ticked in various rows
    For i = 1 To UBound(m_arrintBet)
        For j = i + 1 To UBound(m_arrintBet)
            If m_arrintBet(j) = m_arrintBet(i) Then
                Call ShowInfoPopup(g_c_tool, _
                    GetText(g_arr_Text, "BET029") & " " & m_arrintBet(i) & " " & GetText(g_arr_Text, "BET057"), True, _
                    vbModal, 12)
                m_blnOK = False 'Bet slip is not valid
                Exit For
            End If
        Next
    Next

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "ValidateBetSlip()")
    Call basAuxiliary.CodeCrash
End Sub

'Check if a single row (I, II, III or IV) is filled in correctly
Private Function CheckRow(iRow As Integer, iTicks As Integer) As Integer()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Dim j As Integer
    Dim cnt As Integer, guess(1 To 24) As Integer
    
    cnt = 0
    
    'Loop through the 24 checkboxes of the row and count the ticks
    For j = 1 To 24
        If m_arrintBetSlip(iRow, j) = 1 Then
            cnt = cnt + 1
            guess(cnt) = j 'Note the number of the ticked checkbox
        End If
    Next j
    
    'Check the existance of a horse
    For j = 0 To cnt - 1
        If guess(j + 1) > objRace.NUMBER_ENROLLED Then 'A higher number than enrolled horses is ticked
            Call ShowInfoPopup(g_c_tool, _
                GetText(g_arr_Text, "BET029") & " " & guess(j + 1) & " " & GetText(g_arr_Text, "BET030"), True, _
                vbModal, 12)
            m_blnOK = False 'Bet slip is not valid
        End If
    Next j
    
    'Check the number of ticks in a row
    If cnt < iTicks Then 'Tick is missing
        Call ShowInfoPopup(g_c_tool, _
            GetText(g_arr_Text, "BET031") & " " & WorksheetFunction.Roman(iRow) & " " & GetText(g_arr_Text, "BET032"), True, _
            vbModal, 12)
        m_blnOK = False 'Bet slip is not valid
    ElseIf cnt > iTicks Then 'Too many ticks
        Call ShowInfoPopup(g_c_tool, _
            GetText(g_arr_Text, "BET033") & " " & WorksheetFunction.Roman(iRow), True, _
            vbModal, 12)
        m_blnOK = False 'Bet slip is not valid
    End If
    
    
    
    
    CheckRow = guess() 'Return the number of the ticked checkbox(es)

    Exit Function
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "CheckRow()")
    Call basAuxiliary.CodeCrash
End Function

'Show a pop-up if the minimum stake for this type of bet is violated
Private Sub ErrorMinStake(minStake As Double)
    Call ShowInfoPopup(g_c_tool, _
        GetText(g_arr_Text, "BET034") & " " & minStake & " " & GetText(g_arr_Text, "BET035"), True, _
        vbModal, 12)
    m_blnOK = False
End Sub

Private Function FindHorse(horse As Integer) As Integer
    Dim i As Integer
    For i = 1 To UBound(g_arr_varHorses)
        If g_arr_varHorses(i, 11) = horse Then
            FindHorse = g_arr_varHorses(i, 17) 'Return odds for this horse
            Exit For
        End If
    Next i
End Function

Private Sub opt117_Click() 'Type of bet: Win
    Call ClearBetSlip_II
    Call ClearBetSlip_III
    Call ClearBetSlip_IV
    fraPlace2.Enabled = False
    fraPlace3.Enabled = False
    fraPlace4.Enabled = False
    opt127.Value = True 'Default stake: 2 EUR
End Sub

Private Sub opt118_Click() 'Type of bet: Show
    Call ClearBetSlip_II
    Call ClearBetSlip_III
    Call ClearBetSlip_IV
    fraPlace2.Enabled = False
    fraPlace3.Enabled = False
    fraPlace4.Enabled = False
    opt127.Value = True 'Default stake: 2 EUR
End Sub

Private Sub opt120_Click() 'Type of bet: Exacta
    Call ClearBetSlip_III
    Call ClearBetSlip_IV
    fraPlace2.Enabled = True
    fraPlace3.Enabled = False
    fraPlace4.Enabled = False
    opt126.Value = True 'Default stake: 1 EUR
End Sub

Private Sub opt121_Click() 'Type of bet: 2 sur 4
    Call ClearBetSlip_II
    Call ClearBetSlip_III
    Call ClearBetSlip_IV
    fraPlace2.Enabled = False
    fraPlace3.Enabled = False
    fraPlace4.Enabled = False
    opt128.Value = True 'Default stake: 3 EUR
End Sub

Private Sub opt122_Click() 'Type of bet: Trifecta
    Call ClearBetSlip_IV
    fraPlace2.Enabled = True
    fraPlace3.Enabled = True
    fraPlace4.Enabled = False
    opt125.Value = True 'Default stake: 0.50 EUR
End Sub

Private Sub opt123_Click() 'Type of bet: Superfecta
    fraPlace2.Enabled = True
    fraPlace3.Enabled = True
    fraPlace4.Enabled = True
    opt125.Value = True 'Default stake: 0.50 EUR
End Sub

Private Sub ClearBetSlip_I() 'Clear all Checkboxes in row I
    Dim ctr As control
    For Each ctr In fraPlace1.Controls 'Row I
        ctr.Value = False
    Next ctr
End Sub

Private Sub ClearBetSlip_II() 'Clear all Checkboxes in row II
    Dim ctr As control
    For Each ctr In fraPlace2.Controls 'Row II
        ctr.Value = False
    Next ctr
End Sub

Private Sub ClearBetSlip_III() 'Clear all Checkboxes in row III
    Dim ctr As control
    For Each ctr In fraPlace3.Controls 'Row III
        ctr.Value = False
    Next ctr
End Sub

Private Sub ClearBetSlip_IV() 'Clear all Checkboxes in row IV
    Dim ctr As control
    For Each ctr In fraPlace4.Controls 'Row IV
        ctr.Value = False
    Next ctr
End Sub

'Button DISCARD BET SLIP
Private Sub cmd142_Click()
    Unload Me 'Close the UserForm with the bet slip
End Sub
