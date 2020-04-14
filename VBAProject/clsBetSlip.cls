VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBetSlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class Module for creating object instances for bet slips
'   Class Module clsBetSlip

'OBJECT VARIABLES
Private intID As String 'Bet slip ID
Private strGamblerName As String 'Name of the player
Private dblStake As Double 'Stake (EUR)
Private dblOdds As Double 'Odds (payout for 1 EUR stake)
Private lngType As Long 'Type of the bet
Private strType_Text As String 'Type of the bet (text)
Private intGuess() As Integer 'Guess

'PROPERTIES
'----------
Public Property Get BET_ID() As String
    BET_ID = intID
End Property

Public Property Let BET_ID(ByVal nr As String)
    intID = nr
End Property

Public Property Get GAMBLERNAME() As String
    GAMBLERNAME = strGamblerName
End Property

Public Property Let GAMBLERNAME(ByVal nm As String)
    strGamblerName = nm
End Property

Public Property Get STAKE() As Double
    STAKE = dblStake
End Property

Public Property Let STAKE(ByVal St As Double)
    dblStake = St
End Property

Public Property Get ODDS() As Double
    ODDS = dblOdds
End Property

Public Property Let ODDS(ByVal o As Double)
    dblOdds = o
End Property

Public Property Get BET_TYPE() As Long
    BET_TYPE = lngType
End Property

Public Property Let BET_TYPE(ByVal ty As Long)
    lngType = ty
End Property

Public Property Get BET_TYPE_TEXT() As String
    BET_TYPE_TEXT = strType_Text
End Property

Public Property Let BET_TYPE_TEXT(ByVal t As String)
    strType_Text = t
End Property

Public Property Get BET() As Integer()
    BET = intGuess()
End Property

Public Property Let BET(ByRef b() As Integer)
    intGuess() = b()
End Property
