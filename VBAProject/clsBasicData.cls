VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBasicData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class Module for creating exactly one object instance for storing basic data
'   Class Module clsBasicData

'OBJECT VARIABLES
Private lngExcelStyleNotation As Long 'A1- or R1C1-style notation
Private intTopRows As Integer 'Number of rows at the top of the worksheet (used for the menu in the RS edition)
Private intLeftColumns As Integer 'Number of columns left of the starting gates
Private intColumsAfterFinish As Integer 'Number of columns behind the finish line
Private dblRankingsWidth As Double 'Cell width of the photo of the finish and the ranking list

'PROPERTIES
'----------
Public Property Get XL_STYLE_NOTATION() As Long
    XL_STYLE_NOTATION = lngExcelStyleNotation
End Property

Public Property Let XL_STYLE_NOTATION(ByVal lngX As Long)
    lngExcelStyleNotation = lngX
End Property

Public Property Get TOP_ROWS() As Integer
    TOP_ROWS = intTopRows
End Property

Public Property Let TOP_ROWS(ByVal intX As Integer)
    intTopRows = intX
End Property

Public Property Get LEFT_COLS() As Integer
    LEFT_COLS = intLeftColumns
End Property

Public Property Let LEFT_COLS(ByVal intX As Integer)
    intLeftColumns = intX
End Property

Public Property Get AFTER_FIN_COLS() As Integer
    AFTER_FIN_COLS = intColumsAfterFinish
End Property

Public Property Let AFTER_FIN_COLS(ByVal intX As Integer)
    intColumsAfterFinish = intX
End Property

Public Property Get RANK_WIDTH() As Double
    RANK_WIDTH = dblRankingsWidth
End Property

Public Property Let RANK_WIDTH(ByVal dblX As Double)
    dblRankingsWidth = dblX
End Property
