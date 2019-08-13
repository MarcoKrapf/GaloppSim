VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBetSlip 
   Caption         =   "[Name]"
   ClientHeight    =   6105
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18045
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
Dim m_strType As String 'Type of the bet
Dim m_blnOK As Boolean 'True when bet slip is valid

Private Sub UserForm_Initialize() 'Set default values
    'Section: Race data
    lbl110.caption = GetText(g_arr_Text, "BET005")
    lbl112.caption = GetText(g_arr_Text, "BET003")
    lbl113.caption = GetText(g_arr_Text, "BET004")
    'Section: Checkboxes for place 1-4
    lbl115.caption = g_arr_Grammar(1) & " " & GetText(g_arr_Text, "BET006")
    'Section: Type of bet
    fraBettingType.caption = GetText(g_arr_Text, "BET007")
    opt117.caption = GetText(g_arr_Text, "BET008")
    opt118.caption = GetText(g_arr_Text, "BET009")
    opt120.caption = GetText(g_arr_Text, "BET010")
    opt121.caption = GetText(g_arr_Text, "BET011")
    opt122.caption = GetText(g_arr_Text, "BET012")
    opt123.caption = GetText(g_arr_Text, "BET013")
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
    
    'Default values
    Call opt117_Click   'Type of bet: Win
    opt126.Value = True 'Stake: 1 EUR
    
    'Currently deactivated betting types as they are not implemented yet
    opt118.Enabled = False 'Show
    opt120.Enabled = False 'Exacta
    opt121.Enabled = False 'Place twin
    opt122.Enabled = False 'Trifecta
    opt123.Enabled = False 'Superfecta
    
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
            .id = CStr(g_colBetSlips.count + 1001) & objRace.RACE_ID 'Compile a unique bet slip ID
            .GamblerName = Me.caption
            .Stake = m_dblStake
            .Odds = m_dblOdds
            .BType = m_strType
            .bet = m_arrintBet()
        End With
        
        g_colBetSlips.Add BetSlipN 'Add the bet slip to the Collection
        frmStart.lstBetSlips.AddItem BetSlipN.GamblerName & " - " & Format(BetSlipN.Stake, "0.00") & " " & GetText(g_arr_Text, "BET035") _
                                        & " (" & BetSlipN.BType & ") - [" & GetText(g_arr_Text, "BET001") & " #" & BetSlipN.id & "]" 'Add bet slip to the ListBox
                                        
        Unload Me 'Close the UserForm with the bet slip
        
        objOption.BET_PLACED = True 'Note that at least one bet slip has been submitted
        Call GetNumberBetSlips 'Refresh the number of bet slips
        frmStart.lstBetSlips.Visible = True 'show the area with the bet slips
    End If
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "cmd141_Click()")
    Call basAuxiliary.CodeCrash
End Sub

'Button DISCARD BET SLIP
Private Sub cmd142_Click()
    Unload Me 'Close the UserForm with the bet slip
End Sub

'Read all information filled in by the gambler
Private Sub ReadBetSlip()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Dim ctr As control
    Dim i As Integer
    
    'Get the ticked Checkboxes of each row
    i = 1
    For Each ctr In fraPlace1.Controls 'Top row ("I") for the winner
        m_arrintBetSlip(1, i) = Abs(ctr.Value * 1)
        i = i + 1
        'Comment in the following lines for understanding the possible boolean values of a Checkbox
'            Debug.Print vbNewLine & "Checkbox number " & i - 1 & " value: " & ctr.Value 'True if ticked, False if not
'            Debug.Print "Checkbox value multiplied by 1: " & ctr.Value * 1 '-1 if true, 0 if false
'            Debug.Print "Converted to an absolute value: " & Abs(ctr.Value * 1) '1 if true, 0 if false
    Next ctr

    i = 1
    For Each ctr In fraPlace2.Controls 'Second row ("II") for place 2
        m_arrintBetSlip(2, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr
    
    i = 1
    For Each ctr In fraPlace3.Controls 'Third row ("III") for place 3
        m_arrintBetSlip(3, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr
    
        i = 1
    For Each ctr In fraPlace4.Controls 'Fourth row ("IV") for place 4
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
            m_arrintBet(1) = CheckRow(1) 'Check row I
            #If Debugging Then
                If m_blnOK Then
                    Debug.Print vbNewLine & "Bet on horse number " & m_arrintBet(1) & " - Odds " & FindHorse(m_arrintBet(1)) & ":10"
                Else
                    Debug.Print vbNewLine & "Bet slip not valid"
                End If
            #End If
            m_dblOdds = CDbl(FindHorse(m_arrintBet(1))) / 10 'Calculate the payout for 1 EUR stake
            m_strType = GetText(g_arr_Text, "BET008")
            If opt125 Then Call ErrorMinStake 'If a stake less than 1 EUR is selected
    'NOT IMPLEMENTED YET >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        Case opt118 'Show (not implemented yet)
            m_strType = GetText(g_arr_Text, "BET009")
            If opt125 Then Call ErrorMinStake
        Case opt120 'Exacta (not implemented yet)
            ReDim m_arrintBet(1 To 2)
            m_arrintBet(1) = CheckRow(1) 'Check row I
            m_arrintBet(2) = CheckRow(2) 'Check row II
            m_dblOdds = (CDbl(FindHorse(m_arrintBet(1))) * CDbl(FindHorse(m_arrintBet(2)))) / 10
            m_strType = GetText(g_arr_Text, "BET010")
            If opt125 Then Call ErrorMinStake
        Case opt121 'Place twin (not implemented yet)
            m_strType = GetText(g_arr_Text, "BET011")
            If opt125 Then Call ErrorMinStake
        Case opt122 'Trifecta (not implemented yet)
            ReDim m_arrintBet(1 To 3)
            m_arrintBet(1) = CheckRow(1) 'Check row I
            m_arrintBet(2) = CheckRow(2) 'Check row II
            m_arrintBet(3) = CheckRow(3) 'Check row III
            m_dblOdds = (CDbl(FindHorse(m_arrintBet(1))) * CDbl(FindHorse(m_arrintBet(2))) _
                        * CDbl(FindHorse(m_arrintBet(3))) / 10)
            m_strType = GetText(g_arr_Text, "BET012")
        Case opt123 'Superfecta (not implemented yet)
            ReDim m_arrintBet(1 To 4)
            m_arrintBet(1) = CheckRow(1) 'Check row I
            m_arrintBet(2) = CheckRow(2) 'Check row II
            m_arrintBet(3) = CheckRow(3) 'Check row III
            m_arrintBet(4) = CheckRow(4) 'Check row IV
            m_dblOdds = (CDbl(FindHorse(m_arrintBet(1))) * CDbl(FindHorse(m_arrintBet(2))) _
                        * CDbl(FindHorse(m_arrintBet(3))) * CDbl(FindHorse(m_arrintBet(4))) / 10)
            m_strType = GetText(g_arr_Text, "BET013")
        '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< NOT IMPLEMENTED YET
    End Select

    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "ValidateBetSlip()")
    Call basAuxiliary.CodeCrash
End Sub

'Check if a single row (I, II, III or IV) is filled in correctly
Private Function CheckRow(iRow As Integer) As Integer
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Dim j As Integer
    Dim cnt As Integer, guess As Integer
    
    cnt = 0
    
    'Loop through the 24 Checkboxes of the row
    'and count the ticks
    For j = 1 To 24
        If m_arrintBetSlip(iRow, j) = 1 Then
            cnt = cnt + 1
            guess = j 'Note the number of the ticked Checkbox
        End If
    Next j
    
    If guess > objRace.NUMBER_ENROLLED Then 'A higher number than enrolled horses is ticked
        Call ShowInfoPopup(g_c_tool, _
            GetText(g_arr_Text, "BET029") & " " & guess & " " & GetText(g_arr_Text, "BET030"), True, _
            vbModal)
        m_blnOK = False 'Bet slip is not valid
    ElseIf cnt < 1 Then 'No Checkbox ticked at all
        Call ShowInfoPopup(g_c_tool, _
            GetText(g_arr_Text, "BET031") & " " & iRow & " " & GetText(g_arr_Text, "BET032"), True, _
            vbModal)
        m_blnOK = False 'Bet slip is not valid
    ElseIf cnt > 1 Then 'More than one Checkbox is ticked for this place
        Call ShowInfoPopup(g_c_tool, _
            GetText(g_arr_Text, "BET033") & " " & iRow, True, _
            vbModal)
        m_blnOK = False 'Bet slip is not valid
    Else

    End If
    
    CheckRow = guess 'Return the number of the ticked Checkbox

    Exit Function
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "CheckRow()")
    Call basAuxiliary.CodeCrash
End Function

'Show a pop-up if the minimum stake for this type of bet is violated
Private Sub ErrorMinStake()
    Call ShowInfoPopup(g_c_tool, _
        GetText(g_arr_Text, "BET034") & " " & GetText(g_arr_Text, "BET016") & " " & GetText(g_arr_Text, "BET035"), True, _
        vbModal)
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
    Call ClearBetSlip
    fraPlace2.Enabled = False
    fraPlace3.Enabled = False
    fraPlace4.Enabled = False
End Sub

Private Sub opt120_Click() 'Type of bet: Exacta
    Call ClearBetSlip
    fraPlace2.Enabled = True
    fraPlace3.Enabled = False
    fraPlace4.Enabled = False
End Sub

Private Sub opt122_Click() 'Type of bet: Trifecta
    Call ClearBetSlip
    fraPlace2.Enabled = True
    fraPlace3.Enabled = True
    fraPlace4.Enabled = False
End Sub

Private Sub opt123_Click() 'Type of bet: Superfecta
    Call ClearBetSlip
    fraPlace2.Enabled = True
    fraPlace3.Enabled = True
    fraPlace4.Enabled = True
End Sub

Private Sub ClearBetSlip() 'Clear all Checkboxes
    Dim ctr As control

    For Each ctr In fraPlace1.Controls 'Row I
        ctr.Value = False
    Next ctr
    For Each ctr In fraPlace2.Controls 'Row II
        ctr.Value = False
    Next ctr
    For Each ctr In fraPlace3.Controls 'Row III
        ctr.Value = False
    Next ctr
    For Each ctr In fraPlace4.Controls 'Row IV
        ctr.Value = False
    Next ctr
End Sub
