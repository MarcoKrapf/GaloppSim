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
Private strType As String 'Type of the bet
Private intGuess() As Integer 'Guess

'PROPERTIES
'----------
Public Property Get id() As String
    id = intID
End Property

Public Property Let id(ByVal nr As String)
    intID = nr
End Property

Public Property Get GamblerName() As String
    GamblerName = strGamblerName
End Property

Public Property Let GamblerName(ByVal nm As String)
    strGamblerName = nm
End Property

Public Property Get Stake() As Double
    Stake = dblStake
End Property

Public Property Let Stake(ByVal St As Double)
    dblStake = St
End Property

Public Property Get Odds() As Double
    Odds = dblOdds
End Property

Public Property Let Odds(ByVal o As Double)
    dblOdds = o
End Property

Public Property Get BType() As String
    BType = strType
End Property

Public Property Let BType(ByVal t As String)
    strType = t
End Property

Public Property Get bet() As Integer()
    bet = intGuess()
End Property

Public Property Let bet(ByRef b() As Integer)
    intGuess() = b()
End Property
