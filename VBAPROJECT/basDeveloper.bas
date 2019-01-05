Attribute VB_Name = "basDeveloper"
Option Explicit
Option Private Module

'This module contains non-productive procedures
'which can be used for development purposes

'Execute this procedure for writing the value of the background colour
'of each selected cell into this cell while designing a new race
Sub ReadColourValues()
    Dim r As Range
    For Each r In Selection
        r.Value = r.Interior.color
    Next
End Sub
