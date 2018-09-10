Attribute VB_Name = "basAuxiliary"
Option Explicit
Option Private Module

'Auxiliary procedures

Dim i As Integer

'Place the UserForm in the center of the Window
    'frmMe as UserForm geht nicht: http://www.office-loesung.de/ftopic516038_0_0_asc.php
Public Sub PlaceUserFormInCenter(frmMe As Object)
    With frmMe
        .StartUpPosition = 0
        .top = ActiveWindow.top + ((ActiveWindow.Height - frmMe.Height) / 2)
        .left = ActiveWindow.left + ((ActiveWindow.Width - frmMe.Width) / 2)
    End With
End Sub

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

'Return the language column
Public Function GetLanguageColumn(wks As Worksheet) As Integer
    For i = 1 To 100
        If wks.Cells(1, i).Value = g_strLanguage Then
            GetLanguageColumn = i
            Exit Function
        End If
    Next i
End Function

'Return the text values according to the language
Public Function GetTxt(arr As Variant, ID As String) As String
    Dim r As Long
    For r = 1 To UBound(arr, 2)
        If arr(0, r) = ID Then
            GetTxt = arr(1, r)
            Exit Function
        End If
    Next r
End Function

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
        g_strMsgCaption = g_c_tool
        g_strMsgText = GetTxt(g_arrTxt, "ERROR001")
        'Display the pop-up
        frmMsg_MultiPurpose.Show
End Sub

'Colour picker
Public Function ColPick(current As Long) As Long
    Dim returncode As Integer
    returncode = Application.Dialogs(xlDialogEditColor).Show(10)
    If returncode <> 0 Then
        ColPick = ThisWorkbook.Colors(10)
    Else
        ColPick = current
    End If
End Function

'Formatting: race information on the worksheet
Public Sub RaceInfoWorksheet(colBack As Long, colFore As Long, top)
    'Leader
    With g_wksRace.Cells(1 + top, 2)
        .Font.name = "Arial Black"
        .Font.Size = 8
        .Font.Bold = True
        .Value = ""
    End With
    With g_wksRace.Cells(2 + top, 3)
        .Font.name = "Arial Black"
        .Font.Size = 11
        .Font.Bold = True
        .Value = ""
    End With
    'Race distance progress
    With g_wksRace.Cells(3 + top, 2)
        .Font.name = "Arial Black"
        .Font.Size = 8
        .Value = ""
    End With
    'Colours
    With g_wksRace.Range(Cells(1 + top, 1), Cells(3 + top, 5))
        .Interior.Color = colBack
        .Font.Color = colFore
    End With
End Sub

'Determine the caption of the start button
Public Function getCaptionStartBtn(bettingMode As Boolean) As String
    If bettingMode = False Then
        getCaptionStartBtn = "BTN003a" '"Start the race"
    Else
        getCaptionStartBtn = "BTN003b" '"Betting and race"
    End If
End Function

'Pop-up to delock algorithms
Public Function AllowAlgorithms() As Boolean
    'Set the button mode
    g_strMsgButtons = "YesNo"
    'Assign the text for the pop-up
    g_strMsgCaption = GetTxt(g_arrTxt, "USERFORM007")
    g_strMsgText = GetTxt(g_arrTxt, "ERROR007")
    'Display the pop-up
    frmMsg_MultiPurpose.Show (vbModal) 'modal
    'Evaluate the return value
    If g_strButtonPressed = "YES" Then
        AllowAlgorithms = False
    Else
        AllowAlgorithms = True
    End If
End Function

'Place the cursor far away (in the upper right corner of the screen)
Public Sub CursorAway()
    On Error Resume Next
        g_wksRace.Cells(1, ActiveWindow.VisibleRange.Columns.Count - 1).Activate
End Sub

Public Sub AI_ExcelModeStart()
    Select Case g_strExcelMode
        Case "normal"
            Call ResetExcelOptions
        Case "TVmenu"
            Call ExcelOptionsTVmenu
        Case "TVfull"
            Call ExcelOptionsTVfull
    End Select
End Sub

Public Sub AI_ExcelModeEnd()
    Select Case g_strExcelMode
        Case "normal"
            
        Case "TVmenu"
            
        Case "TVfull"
            Call ExcelOptionsTVmenu
    End Select
End Sub
