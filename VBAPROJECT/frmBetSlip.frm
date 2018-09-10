VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmBetSlip 
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

Dim m_arrintBetSlip(1 To 4, 1 To 24) As Integer 'Entire bet slip
Dim m_arrintBet() As Integer
Dim m_dblStake As Double 'Stake
Dim m_dblOdd As Double 'Odd
Dim m_strType As String 'Type of the bet
Dim m_blnOK As Boolean 'True when bet slip is valid

'Button SHOW HORSE NUMBERS AND ODDS
Private Sub cmd140_Click()
    Call odds
End Sub

'Button PLACE BET
Private Sub cmd141_Click()
    Dim BetSlipN As clsBetSlip
    
    m_blnOK = True
    
    Call ReadBetSlip
    Call ValidateBetSlip
    
    If m_blnOK = True Then 'If bet slip is valid
        Set BetSlipN = New clsBetSlip
        
        With BetSlipN 'Set values
            .ID = CStr(g_colBetSlips.Count + 1001) & g_strRaceID 'compile a unique bet slip ID
            .GamblerName = Me.Caption
            .Stake = m_dblStake
            .Odd = m_dblOdd
            .BType = m_strType
            .bet = m_arrintBet()
        End With
        
        g_colBetSlips.Add BetSlipN 'Add bet slip to Collection
        frmStart.lstBetSlips.AddItem BetSlipN.GamblerName & " - " & Format(BetSlipN.Stake, "0.00") & " " & GetTxt(g_arrTxt, "BET035") _
                                        & " (" & BetSlipN.BType & ") - [" & GetTxt(g_arrTxt, "BET001") & " #" & BetSlipN.ID & "]" 'Add bet slip to ListBox
                                        
        Unload Me
        
        g_blnBetsPlaced = True
        Call NumberBetSlips 'refresh the number of bet slips
        frmStart.lstBetSlips.Visible = True 'show the area with the bet slips
    End If
    
End Sub

'Button DISCARD BET SLIP
Private Sub cmd142_Click()
    Unload Me
End Sub

Private Sub ReadBetSlip()
    Dim ctr As control
    Dim i As Integer
    
    'Checkboxes
    i = 1
    For Each ctr In fraPlace1.Controls
        m_arrintBetSlip(1, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr

    i = 1
    For Each ctr In fraPlace2.Controls
        m_arrintBetSlip(2, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr
    
    i = 1
    For Each ctr In fraPlace3.Controls
        m_arrintBetSlip(3, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr
    
        i = 1
    For Each ctr In fraPlace4.Controls
        m_arrintBetSlip(4, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr
    
    'Stake
    Select Case True
        Case opt125
            m_dblStake = CDbl(GetTxt(g_arrTxt, "BET015"))
        Case opt126
            m_dblStake = CDbl(GetTxt(g_arrTxt, "BET016"))
        Case opt127
            m_dblStake = CDbl(GetTxt(g_arrTxt, "BET017"))
        Case opt128
            m_dblStake = CDbl(GetTxt(g_arrTxt, "BET018"))
        Case opt129
            m_dblStake = CDbl(GetTxt(g_arrTxt, "BET019"))
        Case opt130
            m_dblStake = CDbl(GetTxt(g_arrTxt, "BET020"))
        Case opt131
            m_dblStake = CDbl(GetTxt(g_arrTxt, "BET021"))
        Case opt132
            m_dblStake = CDbl(GetTxt(g_arrTxt, "BET022"))
        Case opt133
            m_dblStake = CDbl(GetTxt(g_arrTxt, "BET023"))
        Case opt134
            m_dblStake = CDbl(GetTxt(g_arrTxt, "BET024"))
        Case opt135
            m_dblStake = CDbl(GetTxt(g_arrTxt, "BET025"))
    End Select
End Sub

Private Sub ValidateBetSlip()
    Select Case True
        Case opt117 'Win
            ReDim m_arrintBet(1 To 1)
            m_arrintBet(1) = CheckRow(1)
            m_dblOdd = CDbl(FindHorse(m_arrintBet(1))) / 10
            m_strType = GetTxt(g_arrTxt, "BET008")
            If opt125 Then Call ErrorMinStake
        Case opt118 'Place???Show???
            m_strType = GetTxt(g_arrTxt, "BET009")
            If opt125 Then Call ErrorMinStake
        Case opt120 'Exacta
            ReDim m_arrintBet(1 To 2)
            m_arrintBet(1) = CheckRow(1)
            m_arrintBet(2) = CheckRow(2)
            m_dblOdd = (CDbl(FindHorse(m_arrintBet(1))) * CDbl(FindHorse(m_arrintBet(2)))) / 10
            m_strType = GetTxt(g_arrTxt, "BET010")
            Debug.Print "Tipp auf Startnummer " & m_arrintBet(1) & " und " & m_arrintBet(2) & " - Quote " & m_dblOdd
            If opt125 Then Call ErrorMinStake
        Case opt121 'PZW????????
            m_strType = GetTxt(g_arrTxt, "BET011")
            If opt125 Then Call ErrorMinStake
        Case opt122 'Trifecta
            ReDim m_arrintBet(1 To 3)
            m_arrintBet(1) = CheckRow(1)
            m_arrintBet(2) = CheckRow(2)
            m_arrintBet(3) = CheckRow(3)
            m_dblOdd = (CDbl(FindHorse(m_arrintBet(1))) * CDbl(FindHorse(m_arrintBet(2))) _
                        * CDbl(FindHorse(m_arrintBet(3))) / 10)
            m_strType = GetTxt(g_arrTxt, "BET012")
            Debug.Print "Tipp auf Startnummer " & m_arrintBet(1) & " und " & m_arrintBet(2) & " und " & m_arrintBet(3) & " - Quote " & m_dblOdd
        Case opt123 'Superfecta
            ReDim m_arrintBet(1 To 4)
            m_arrintBet(1) = CheckRow(1)
            m_arrintBet(2) = CheckRow(2)
            m_arrintBet(3) = CheckRow(3)
            m_arrintBet(4) = CheckRow(4)
            m_dblOdd = (CDbl(FindHorse(m_arrintBet(1))) * CDbl(FindHorse(m_arrintBet(2))) _
                        * CDbl(FindHorse(m_arrintBet(3))) * CDbl(FindHorse(m_arrintBet(4))) / 10)
            m_strType = GetTxt(g_arrTxt, "BET013")
            Debug.Print "Tipp auf Startnummer " & m_arrintBet(1) & " und " & m_arrintBet(2) & " und " & m_arrintBet(3) & " und " & m_arrintBet(4) & _
                " - Quote " & m_dblOdd
    End Select
End Sub

Private Function CheckRow(i As Integer) As Integer
    Dim j As Integer
    Dim cnt As Integer, guess As Integer
    
    cnt = 0
    
    For j = 1 To 24
        If m_arrintBetSlip(i, j) = 1 Then
            cnt = cnt + 1
            guess = j
        End If
    Next j
    
    If guess > g_intHorsesEnrolled Then
        'Pop-up
            'Set the button mode
            g_strMsgButtons = "OK"
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_tool
            g_strMsgText = GetTxt(g_arrTxt, "BET029") & " " & guess & " " & GetTxt(g_arrTxt, "BET030") & "."
            'Display the pop-up
            frmMsg_Attention.Show
        m_blnOK = False
    ElseIf cnt < 1 Then
        'Pop-up
            'Set the button mode
            g_strMsgButtons = "OK"
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_tool
            g_strMsgText = GetTxt(g_arrTxt, "BET031") & " " & i & " " & GetTxt(g_arrTxt, "BET032")
            'Display the pop-up
            frmMsg_Attention.Show
        m_blnOK = False
    ElseIf cnt > 1 Then
        'Pop-up
            'Set the button mode
            g_strMsgButtons = "OK"
            'Assign the text for the pop-up
            g_strMsgCaption = g_c_tool
            g_strMsgText = GetTxt(g_arrTxt, "BET033") & " " & i
            'Display the pop-up
            frmMsg_Attention.Show
        m_blnOK = False
    Else

    End If
    
    CheckRow = guess

End Function

Private Sub ErrorMinStake()
    'Pop-up
        'Set the button mode
        g_strMsgButtons = "OK"
        'Assign the text for the pop-up
        g_strMsgCaption = g_c_tool
        g_strMsgText = GetTxt(g_arrTxt, "BET034") & " " & GetTxt(g_arrTxt, "BET016") & " " & GetTxt(g_arrTxt, "BET035")
        'Display the pop-up
        frmMsg_Attention.Show
    m_blnOK = False
End Sub

Private Function FindHorse(horse As Integer) As Integer
    Dim i As Integer
    For i = 1 To UBound(g_arr_varHorses)
        If g_arr_varHorses(i, 11) = horse Then
            FindHorse = g_arr_varHorses(i, 17) 'Return odd
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

Private Sub ClearBetSlip() 'Clear checkboxes
    Dim ctr As control

    For Each ctr In fraPlace1.Controls
        ctr.Value = False
    Next ctr
    For Each ctr In fraPlace2.Controls
        ctr.Value = False
    Next ctr
    For Each ctr In fraPlace3.Controls
        ctr.Value = False
    Next ctr
    For Each ctr In fraPlace4.Controls
        ctr.Value = False
    Next ctr
End Sub

Private Sub UserForm_Initialize() 'Default values
    'Section: Race data
    lbl110.Caption = GetTxt(g_arrTxt, "BET005")
    lbl112.Caption = GetTxt(g_arrTxt, "BET003")
    lbl113.Caption = GetTxt(g_arrTxt, "BET004")
    'Section: Checkboxes
    lbl115.Caption = GetTxt(g_arrTxt, "BET006")
    'Section: Type of bet
    fraBettingType.Caption = GetTxt(g_arrTxt, "BET007")
    opt117.Caption = GetTxt(g_arrTxt, "BET008")
    opt118.Caption = GetTxt(g_arrTxt, "BET009")
    opt120.Caption = GetTxt(g_arrTxt, "BET010")
    opt121.Caption = GetTxt(g_arrTxt, "BET011")
    opt122.Caption = GetTxt(g_arrTxt, "BET012")
    opt123.Caption = GetTxt(g_arrTxt, "BET013")
    'Section: Stake
    fraStake.Caption = GetTxt(g_arrTxt, "BET014")
    opt125.Caption = GetTxt(g_arrTxt, "BET015")
    opt126.Caption = GetTxt(g_arrTxt, "BET016")
    opt127.Caption = GetTxt(g_arrTxt, "BET017")
    opt128.Caption = GetTxt(g_arrTxt, "BET018")
    opt129.Caption = GetTxt(g_arrTxt, "BET019")
    opt130.Caption = GetTxt(g_arrTxt, "BET020")
    opt131.Caption = GetTxt(g_arrTxt, "BET021")
    opt132.Caption = GetTxt(g_arrTxt, "BET022")
    opt133.Caption = GetTxt(g_arrTxt, "BET023")
    opt134.Caption = GetTxt(g_arrTxt, "BET024")
    opt135.Caption = GetTxt(g_arrTxt, "BET025")
    'Buttons
    cmd140.Caption = GetTxt(g_arrTxt, "START008") '"Show horse numbers and odds"
    cmd141.Caption = GetTxt(g_arrTxt, "BET026") '"Place bet"
    cmd142.Caption = GetTxt(g_arrTxt, "BET027") '"Discard betting slip"
    
    'Default values
    Call opt117_Click   'Type of bet: Win
    opt126.Value = True 'Stake: 1 EUR
    
    'Currently deactivated
    opt118.Enabled = False 'Platz
    opt120.Enabled = False 'Zweier
    opt121.Enabled = False 'PZW
    opt122.Enabled = False 'Dreier
    opt123.Enabled = False 'Vierer
    
    'Display the UserForm in the center of the Window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub
