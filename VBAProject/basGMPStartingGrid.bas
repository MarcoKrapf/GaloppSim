Attribute VB_Name = "basGMPStartingGrid"
Option Explicit

'This module contains a procedure for drawing
'starting grids for GMP semi finals and the final race
'   Module basGMPStartingGrid

Private Sub GMP_DrawStartingGrid()

    Dim wks As Worksheet
    Set wks = ActiveSheet
    
    Dim col1 As Long
    Dim col2 As Long
    Dim col3 As Long
    
    Dim pool1 As Integer
    Dim pool2 As Integer
    Dim pool3 As Integer
    
    Dim endOfAllocation As Integer
    
    Dim i As Integer, j As Integer
    
    col1 = wks.Cells(6, 2).Interior.Color
    col2 = wks.Cells(6, 4).Interior.Color
    col3 = wks.Cells(6, 6).Interior.Color
    
    pool1 = wks.Cells(rows.count, 2).End(xlUp).row - 5
    pool2 = wks.Cells(rows.count, 4).End(xlUp).row - 5
    pool3 = wks.Cells(rows.count, 6).End(xlUp).row - 5
    
    endOfAllocation = wks.Cells(rows.count, 9).End(xlUp).row
    
    Pause
    'Pool 1
    For i = 6 To pool1 + 5
        With wks.Cells(i, 2)
            .BorderAround Weight:=xlThin
            .Interior.Color = vbGreen
        End With

        Pause
        Do
            j = Int((endOfAllocation - 6 + 1) * Rnd + 6)
            If wks.Cells(j, 10).Interior.Color = col1 And wks.Cells(j, 10).Value = "" Then
                With wks.Cells(j, 10)
                    .Value = wks.Cells(i, 2).Value
                    .Interior.Color = vbWhite
                    .Font.Bold = True
                    .BorderAround Weight:=xlThin
                    .Interior.Color = vbGreen
                End With
                wks.Cells(i, 2).Clear
                Pause2
                With wks.Cells(j, 10)
                    .Borders.LineStyle = xlNone
                    .Interior.Color = vbWhite
                End With
                Exit Do
            End If
        Loop
    Next
    
    'Pool 2
    For i = 6 To pool2 + 5
        With wks.Cells(i, 4)
            .BorderAround Weight:=xlThin
            .Interior.Color = vbGreen
        End With
        Pause
        Do
            j = Int((endOfAllocation - 6 + 1) * Rnd + 6)
            If wks.Cells(j, 10).Interior.Color = col2 And wks.Cells(j, 10).Value = "" Then
                With wks.Cells(j, 10)
                    .Value = wks.Cells(i, 4).Value
                    .Interior.Color = vbWhite
                    .Font.Bold = True
                    .BorderAround Weight:=xlThin
                    .Interior.Color = vbGreen
                End With
                wks.Cells(i, 4).Clear
                Pause2
                With wks.Cells(j, 10)
                    .Borders.LineStyle = xlNone
                    .Interior.Color = vbWhite
                End With
                Exit Do
            End If
        Loop
    Next
    
    'Pool 3
    For i = 6 To pool3 + 5
        With wks.Cells(i, 6)
            .BorderAround Weight:=xlThin
            .Interior.Color = vbGreen
        End With
        Pause
        Do
            j = Int((endOfAllocation - 6 + 1) * Rnd + 6)
            If wks.Cells(j, 10).Interior.Color = col3 And wks.Cells(j, 10).Value = "" Then
                With wks.Cells(j, 10)
                    .Value = wks.Cells(i, 6).Value
                    .Interior.Color = vbWhite
                    .Font.Bold = True
                    .BorderAround Weight:=xlThin
                    .Interior.Color = vbGreen
                End With
                wks.Cells(i, 6).Clear
                Pause2
                With wks.Cells(j, 10)
                    .Borders.LineStyle = xlNone
                    .Interior.Color = vbWhite
                End With
                Exit Do
            End If
        Loop
    Next

End Sub

Private Sub Pause()
    Application.Wait (Now + TimeValue("0:00:05")) '5
End Sub

Private Sub Pause2()
    Application.Wait (Now + TimeValue("0:00:02")) '2
End Sub

