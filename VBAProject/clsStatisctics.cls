VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStatisctics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class Module for creating exactly one object instance for storing statistical values
'   Class Module clsStatisctics

'OBJECT VARIABLES
'----------------
Private intLeaderNr As Integer 'Horse number of the current leader
Private strLeaderName As String 'Horse name of the current leader
Private lngLeaderPos As Long 'Position of the current leader
Private strWinnerName As String 'Name(s) of the horse(s) at 1st place

'PROPERTIES
'----------
Public Property Get LEADER_NR() As Integer
    LEADER_NR = intLeaderNr
End Property

Public Property Let LEADER_NR(ByVal intX As Integer)
    intLeaderNr = intX
End Property

Public Property Get LEADER_NAME() As String
    LEADER_NAME = strLeaderName
End Property

Public Property Let LEADER_NAME(ByVal strX As String)
    strLeaderName = strX
End Property

Public Property Get LEADER_POSITION() As Long
    LEADER_POSITION = lngLeaderPos
End Property

Public Property Let LEADER_POSITION(ByVal lngX As Long)
    lngLeaderPos = lngX
End Property

Public Property Get WINNER_NAME() As String
    WINNER_NAME = strWinnerName
End Property

Public Property Let WINNER_NAME(ByVal strX As String)
    strWinnerName = strX
End Property
