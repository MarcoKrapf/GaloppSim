VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class Module for creating exactly one object instance for storing speed values
'   Class Module clsSpeed

'OBJECT VARIABLES
'----------------
'Range of the basic speed in case it is not fixed for a horse
    Private lngSpeedBasicLow As Long 'Standard value: 1480
    Private lngSpeedBasicHigh As Long 'Standard value: 1520
'Range of the daily form of the horses
    Private lngSpeedCondLow As Long 'Standard value: 1490
    Private lngSpeedCondHigh As Long 'Standard value: 1510
'Range of the randomly assigned speed per step
    Private lngSpeedLoopLow As Long 'Standard value: 0
    Private lngSpeedLoopHigh As Long 'Standard value: 3000
'Range of each phase if racing tactics are active
    Private lngSpeedTacticsLow As Long 'Standard value: 1200
    Private lngSpeedTacticsMedium As Long 'Standard value: 1500
    Private lngSpeedTacticsHigh As Long 'Standard value: 1800


'PROPERTIES
'----------
Public Property Get SPEED_BASIC_LOW() As Long
    SPEED_BASIC_LOW = lngSpeedBasicLow
End Property

Public Property Let SPEED_BASIC_LOW(ByVal lngX As Long)
    lngSpeedBasicLow = lngX
End Property

Public Property Get SPEED_BASIC_HIGH() As Long
    SPEED_BASIC_HIGH = lngSpeedBasicHigh
End Property

Public Property Let SPEED_BASIC_HIGH(ByVal lngX As Long)
    lngSpeedBasicHigh = lngX
End Property



Public Property Get SPEED_COND_LOW() As Long
    SPEED_COND_LOW = lngSpeedCondLow
End Property

Public Property Let SPEED_COND_LOW(ByVal lngX As Long)
    lngSpeedCondLow = lngX
End Property

Public Property Get SPEED_COND_HIGH() As Long
    SPEED_COND_HIGH = lngSpeedCondHigh
End Property

Public Property Let SPEED_COND_HIGH(ByVal lngX As Long)
    lngSpeedCondHigh = lngX
End Property



Public Property Get SPEED_LOOP_LOW() As Long
    SPEED_LOOP_LOW = lngSpeedLoopLow
End Property

Public Property Let SPEED_LOOP_LOW(ByVal lngX As Long)
    lngSpeedLoopLow = lngX
End Property

Public Property Get SPEED_LOOP_HIGH() As Long
    SPEED_LOOP_HIGH = lngSpeedLoopHigh
End Property

Public Property Let SPEED_LOOP_HIGH(ByVal lngX As Long)
    lngSpeedLoopHigh = lngX
End Property



Public Property Get SPEED_TACTICS_LOW() As Long
    SPEED_TACTICS_LOW = lngSpeedTacticsLow
End Property

Public Property Let SPEED_TACTICS_LOW(ByVal lngX As Long)
    lngSpeedTacticsLow = lngX
End Property

Public Property Get SPEED_TACTICS_MEDIUM() As Long
    SPEED_TACTICS_MEDIUM = lngSpeedTacticsMedium
End Property

Public Property Let SPEED_TACTICS_MEDIUM(ByVal lngX As Long)
    lngSpeedTacticsMedium = lngX
End Property

Public Property Get SPEED_TACTICS_HIGH() As Long
    SPEED_TACTICS_HIGH = lngSpeedTacticsHigh
End Property

Public Property Let SPEED_TACTICS_HIGH(ByVal lngX As Long)
    lngSpeedTacticsHigh = lngX
End Property
