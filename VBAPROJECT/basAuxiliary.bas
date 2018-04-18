Attribute VB_Name = "basAuxiliary"
Option Explicit
Option Private Module

'Auxiliary procedures 'ToDo sauber beschreiben.........

Dim i As Integer

'Freezing/unfreezing the window pane
Public Sub Freeze(col As Integer, row As Integer, direction As Boolean)
    With Application.ActiveWindow
        .SplitColumn = col
        .SplitRow = row
        .FreezePanes = direction
    End With
End Sub

'Scrolling
Public Sub Scroll(col As Integer, row As Integer)
    On Error Resume Next
        With Application.ActiveWindow
            .ScrollColumn = col
            .ScrollRow = row
        End With
    On Error GoTo 0
End Sub

'Return the column for a picture from the worksheet "Pic"
Public Function GetPictureColumn(strPic As String) As Integer
    For i = 1 To g_wksPicData.Cells(1, Columns.Count).End(xlToLeft).Column
        If g_wksPicData.Cells(1, i).Value = strPic Then Exit For 'exit loop if the column is found
    Next i
    GetPictureColumn = i 'return the column number
End Function

'Pop-up showing an error message
Public Sub CodeCrash()
    'Pop-up
        'Set the button mode
        g_strMsgButtons = "OK"
        'Assign the text for the pop-up
        g_strMsgCaption = g_c_TOOL
        g_strMsgText = g_strTxt(99)
        'Display the pop-up
        frmMsg_MultiPurpose.Show
End Sub
