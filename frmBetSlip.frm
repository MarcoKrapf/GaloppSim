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

Dim intBetSlip(1 To 4, 1 To 24) As Integer 'Entire bet slip
Dim intBet() As Integer
Dim dblStake As Double 'Stake
Dim dblOdd As Double 'Odd
Dim strType As String 'Type of the bet
Dim blnOK As Boolean 'True when bet slip is valid

'Button PLACE BET
Private Sub cmd141_Click()
    Dim BetSlipN As clsBetSlip
    
    blnOK = True
    
    Call ReadBetSlip
    Call ValidateBetSlip
    
    If blnOK = True Then 'If bet slip is valid
        Set BetSlipN = New clsBetSlip
        
        With BetSlipN 'Set values
            .id = "04020535HH" & CStr(collSlips.Count + 1001)
            .GamblerName = Me.Caption
            .Stake = dblStake
            .Odd = dblOdd
            .BType = strType
            .bet = intBet()
        End With
        
        collSlips.Add BetSlipN 'Add bet slip to Collection
        frmStart.boxBetSlips.AddItem BetSlipN.id & " - " & BetSlipN.GamblerName _
            & " - " & Format(BetSlipN.Stake, "0.00") & " " & txt(151) 'Add bet slip to ListBox
        
        Unload Me
        
        betting = True
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
    For Each ctr In frmP1.Controls
        intBetSlip(1, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr

    i = 1
    For Each ctr In frmP2.Controls
        intBetSlip(2, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr
    
    i = 1
    For Each ctr In frmP3.Controls
        intBetSlip(3, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr
    
        i = 1
    For Each ctr In frmP4.Controls
        intBetSlip(4, i) = Abs(ctr.Value * 1)
        i = i + 1
    Next ctr
    
    'Stake
    Select Case True
        Case opt125
            dblStake = CDbl(txt(125))
        Case opt126
            dblStake = CDbl(txt(126))
        Case opt127
            dblStake = CDbl(txt(127))
        Case opt128
            dblStake = CDbl(txt(128))
        Case opt129
            dblStake = CDbl(txt(129))
        Case opt130
            dblStake = CDbl(txt(130))
        Case opt131
            dblStake = CDbl(txt(131))
        Case opt132
            dblStake = CDbl(txt(132))
        Case opt133
            dblStake = CDbl(txt(133))
        Case opt134
            dblStake = CDbl(txt(134))
        Case opt135
            dblStake = CDbl(txt(135))
    End Select
End Sub

Private Sub ValidateBetSlip()
    Select Case True
        Case opt117 'Win
            ReDim intBet(1 To 1)
            intBet(1) = CheckRow(1)
'            Debug.Print "Tipp auf Startnummer " & intBet(1) & " - Quote " & FindHorse(intBet(1))
            dblOdd = CDbl(FindHorse(intBet(1))) / 10
            strType = txt(117)
            If opt125 Then Call ErrorMinStake
        Case opt118 'Place???Show???
            strType = txt(118)
            If opt125 Then Call ErrorMinStake
        Case opt120 'Exacta
            ReDim intBet(1 To 2)
            intBet(1) = CheckRow(1)
            intBet(2) = CheckRow(2)
            dblOdd = (CDbl(FindHorse(intBet(1))) * CDbl(FindHorse(intBet(2)))) / 10
            strType = txt(120)
            Debug.Print "Tipp auf Startnummer " & intBet(1) & " und " & intBet(2) & " - Quote " & dblOdd
            If opt125 Then Call ErrorMinStake
        Case opt121 'PZW????????
            strType = txt(121)
            If opt125 Then Call ErrorMinStake
        Case opt122 'Trifecta
            ReDim intBet(1 To 3)
            intBet(1) = CheckRow(1)
            intBet(2) = CheckRow(2)
            intBet(3) = CheckRow(3)
            dblOdd = (CDbl(FindHorse(intBet(1))) * CDbl(FindHorse(intBet(2))) _
                        * CDbl(FindHorse(intBet(3))) / 10)
            strType = txt(122)
            Debug.Print "Tipp auf Startnummer " & intBet(1) & " und " & intBet(2) & " und " & intBet(3) & " - Quote " & dblOdd
        Case opt123 'Superfecta
            ReDim intBet(1 To 4)
            intBet(1) = CheckRow(1)
            intBet(2) = CheckRow(2)
            intBet(3) = CheckRow(3)
            intBet(4) = CheckRow(4)
            dblOdd = (CDbl(FindHorse(intBet(1))) * CDbl(FindHorse(intBet(2))) _
                        * CDbl(FindHorse(intBet(3))) * CDbl(FindHorse(intBet(4))) / 10)
            strType = txt(123)
            Debug.Print "Tipp auf Startnummer " & intBet(1) & " und " & intBet(2) & " und " & intBet(3) & " und " & intBet(4) & _
                " - Quote " & dblOdd
    End Select
End Sub

Private Function CheckRow(i As Integer) As Integer
    Dim j As Integer
    Dim cnt As Integer, guess As Integer
    
    cnt = 0
    
    For j = 1 To 24
        If intBetSlip(i, j) = 1 Then
            cnt = cnt + 1
            guess = j
        End If
    Next j
    
    If cnt < 1 Then
        MsgBox txt(146) & " " & i & " " & txt(147), , ""
        blnOK = False
    ElseIf cnt > 1 Then
        MsgBox txt(148) & " " & txt(146) & " " & i & txt(147), , ""
        blnOK = False
    Else

    End If
    
    CheckRow = guess

End Function

Private Sub ErrorMinStake()
    MsgBox txt(150) & " " & txt(126) & " " & txt(151), , ""
    blnOK = False
End Sub

Private Function FindHorse(horse As Integer) As Integer
    Dim i As Integer
    For i = 1 To UBound(arrayPferde)
        If arrayPferde(i, 11) = horse Then
            FindHorse = arrayPferde(i, 17) 'Return odd
            Exit For
        End If
    Next i
End Function

Private Sub opt117_Click() 'Type of bet: Win
    Call ClearBetSlip
    frmP2.Enabled = False
    frmP3.Enabled = False
    frmP4.Enabled = False
End Sub

Private Sub opt120_Click() 'Type of bet: Exacta
    Call ClearBetSlip
    frmP2.Enabled = True
    frmP3.Enabled = False
    frmP4.Enabled = False
End Sub

Private Sub opt122_Click() 'Type of bet: Trifecta
    Call ClearBetSlip
    frmP2.Enabled = True
    frmP3.Enabled = True
    frmP4.Enabled = False
End Sub

Private Sub opt123_Click() 'Type of bet: Superfecta
    Call ClearBetSlip
    frmP2.Enabled = True
    frmP3.Enabled = True
    frmP4.Enabled = True
End Sub

Private Sub ClearBetSlip() 'Clear checkboxes
    Dim ctr As control

    For Each ctr In frmP1.Controls
        ctr.Value = False
    Next ctr
    For Each ctr In frmP2.Controls
        ctr.Value = False
    Next ctr
    For Each ctr In frmP3.Controls
        ctr.Value = False
    Next ctr
    For Each ctr In frmP4.Controls
        ctr.Value = False
    Next ctr
End Sub

Private Sub UserForm_Initialize() 'Default values
    'Captions
    lbl110.Caption = txt(114)
    lbl112.Caption = txt(112)
    lbl113.Caption = txt(113)
    frm116.Caption = txt(116)
    opt117.Caption = txt(117)
    opt118.Caption = txt(118)
    opt120.Caption = txt(120)
    opt121.Caption = txt(121)
    opt122.Caption = txt(122)
    opt123.Caption = txt(123)
    frm124.Caption = txt(124)
    opt125.Caption = txt(125)
    opt126.Caption = txt(126)
    opt127.Caption = txt(127)
    opt128.Caption = txt(128)
    opt129.Caption = txt(129)
    opt130.Caption = txt(130)
    opt131.Caption = txt(131)
    opt132.Caption = txt(132)
    opt133.Caption = txt(133)
    opt134.Caption = txt(134)
    opt135.Caption = txt(135)
    cmd141.Caption = txt(141)
    cmd142.Caption = txt(142)
    
    Call opt117_Click   'Type of bet: Win
    opt126.Value = True 'Stake: 1 EUR
    
    'Currently deactivated
    opt118.Enabled = False 'Platz
    opt120.Enabled = False 'Zweier
    opt121.Enabled = False 'PZW
    opt122.Enabled = False 'Dreier
    opt123.Enabled = False 'Vierer
End Sub
