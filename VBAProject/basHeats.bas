Attribute VB_Name = "basHeats"
Option Explicit

'This module contains a procedure for drawing heats
'   Module DrawHeats

Private Sub DrawHeats()
    Dim wks As Worksheet
    Set wks = ActiveSheet
    
    Dim firstRow As Integer, lastRow As Integer
    Dim i As Integer, j As Integer
    
    firstRow = 4
    lastRow = wks.Cells(rows.count, 2).End(xlUp).row
    
    Pause
    
    For i = firstRow To lastRow
        With wks.Cells(i, 2)
            .BorderAround Weight:=xlThin
            .Interior.Color = vbGreen
        End With

        Pause
        Do
            j = Int((lastRow - firstRow + 1) * Rnd + firstRow)
            If wks.Cells(j, 7).Value = "" Then
                With wks.Cells(j, 7)
                    .Value = wks.Cells(i, 2).Value
                    .Font.Bold = True
                    .Interior.Color = vbGreen
                End With
                wks.Cells(i, 2).Clear
                Pause2
                With wks.Cells(j, 7)
                    .Interior.Color = wks.Cells(j, 6).Interior.Color
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
