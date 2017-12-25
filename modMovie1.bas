Attribute VB_Name = "modMovie1"
Option Explicit

Dim wksMovie As Worksheet, wksPic As Worksheet

Public Sub movie1(txt() As String, ActSheet As String)
    
    Dim wkscheck As Worksheet
    
    'Check whether WorkSheet exists
    For Each wkscheck In ActiveWorkbook.Worksheets
        If wkscheck.name = "GALOPPSIM_MOVIE" Then 'WorkSheet exists
            Application.DisplayAlerts = False
            wkscheck.Delete 'delete WorkSheet
            Application.DisplayAlerts = True
        End If
    Next wkscheck
    'Create WorkSheet
        Set wksMovie = ActiveWorkbook.Worksheets.Add(before:=Sheets(1))
        With wksMovie
            .name = "GALOPPSIM_MOVIE"
            .Range(Columns(1), Columns(100)).ColumnWidth = 2
            .Activate
        End With
        
    Set wksPic = ThisWorkbook.Worksheets("Pic")
    
    wksMovie.Activate
    wksMovie.Cells(4, 5).Font.name = "MV Boli"
    wksMovie.Cells(5, 8).Font.name = "MV Boli"
    wksMovie.Range(Cells(2, 38), Cells(3, 38)).Font.name = "Arial Rounded MT Bold"
    wksMovie.Range(Cells(2, 28), Cells(5, 51)).Font.name = "Arial Rounded MT Bold"
    wksMovie.Range(Cells(2, 22), Cells(4, 28)).Font.name = "Arial Rounded MT Bold"
    With wksMovie.Cells(20, 59)
        .Font.FontStyle = "Italic"
        .Value = txt(400)
    End With
    
    Call DrawPicture(5)
    Call Wait("0:00:02")
    Call Opening("0:00:04", "0:00:02")
    Call Speaker("0:00:02", "0:00:02", txt(401))
    
    Call DrawPicture(6)
    Call BetSlip1
    Call Krapf("0:00:04", "0:00:02", txt(402))
    Call BetSlip2
    
    Call DrawPicture(5)
    Call Wait("0:00:02")
    Call Speaker("0:00:02", "0:00:02", txt(403))
    Call Speaker("0:00:04", "0:00:02", txt(404), txt(405))
    Call Speaker("0:00:04", "0:00:00", txt(406), txt(407))
    Call Speaker("0:00:04", "0:00:00", txt(408))
    Call Speaker("0:00:04", "0:00:00", txt(409))
    
    Call DrawPicture(7)
    Call Speaker("0:00:04", "0:00:00", txt(410), txt(411))
    
    Call DrawPicture(8)
    Call Speaker("0:00:04", "0:00:00", txt(412), txt(413))
    
    Call DrawPicture(9)
    Call Speaker("0:00:04", "0:00:00", txt(414), txt(415))
    
    Call DrawPicture(10)
    Call Speaker("0:00:04", "0:00:00", txt(416))
    
    Call DrawPicture(11)
    Call Speaker("0:00:04", "0:00:00", txt(417), txt(418))
    
    Call DrawPicture(5)
    Call Speaker("0:00:04", "0:00:00", txt(419))
    Call Speaker("0:00:04", "0:00:00", txt(420), txt(421))
    Call Speaker("0:00:04", "0:00:00", txt(422))
    Call Speaker("0:00:04", "0:00:00", txt(423), txt(424))
    Call Speaker("0:00:04", "0:00:00", txt(425), txt(426))
    Call Speaker("0:00:04", "0:00:00", txt(427), txt(428))
    
    Call DrawPicture(12)
    Call Speaker("0:00:04", "0:00:00", txt(429))
    
    Call DrawPicture(13)
    Call Speaker("0:00:04", "0:00:00", txt(430), txt(431))
    
    Call DrawPicture(14)
    Call Speaker("0:00:04", "0:00:00", txt(432), txt(433))
    
    Call DrawPicture(15)
    wksMovie.Range(Cells(1, 5), Cells(7, 8)).Font.FontStyle = "Bold"
    Call Speaker("0:00:04", "0:00:00", txt(434))
    
    Call DrawPicture(16)
    Call Speaker("0:00:04", "0:00:00", txt(435), txt(436))
    
    Call DrawPicture(17)
    Call Speaker("0:00:04", "0:00:00", txt(437), txt(438))
    wksMovie.Range(Cells(1, 5), Cells(7, 8)).Font.FontStyle = "Regular"
    Call Speaker("0:00:04", "0:00:02", txt(439))
    
    Call DrawPicture(18)
    Call Wait("0:00:02")
    Call Krapf("0:00:04", "0:00:02", txt(440))
    Call Leuerer("0:00:04", "0:00:02", txt(441))
    
    Call DrawPicture(19)
    Call Wait("0:00:02")
    
    Call DrawPicture(20)
    Call Wait("0:00:02")
    
    Call DrawPicture(21)
    Call Wait("0:00:04")
    
    ActiveWorkbook.Sheets(ActSheet).Activate 'Go back to WorkSheet
    Application.DisplayAlerts = False
    wksMovie.Delete 'Delete movie WorkSheet
    Application.DisplayAlerts = True

End Sub

Private Sub DrawPicture(PicColumn As Integer)

Dim i As Integer, j As Integer
Dim PicRow As Integer

    PicRow = 2
    
    Application.ScreenUpdating = False
    
    For i = 1 To 40
        For j = 1 To 100
            wksMovie.Cells(i, j).Interior.Color = wksPic.Cells(PicRow, PicColumn).Value
            PicRow = PicRow + 1
        Next
    Next
    
    wksMovie.Cells(41, 1).Select
    Application.ScreenUpdating = True

End Sub

Private Sub WriteText(r As Integer, c As Integer, t As String)
    wksMovie.Cells(r, c).Value = t
End Sub

Private Sub Speaker(wait1 As String, wait2 As String, t As String, Optional t2 As String)

    Application.ScreenUpdating = False
    Call WriteText(3, 6, "/")
    Call WriteText(3, 7, "/")
    Call WriteText(2, 7, "/")
    Call WriteText(2, 8, "/")
    Call WriteText(1, 8, "/")
    Call WriteText(5, 6, "\")
    Call WriteText(5, 7, "\")
    Call WriteText(6, 7, "\")
    Call WriteText(6, 8, "\")
    Call WriteText(7, 8, "\")
    Call WriteText(4, 5, t)
    Call WriteText(5, 8, t2)
    Application.ScreenUpdating = True
    Application.Wait (Now + TimeValue(wait1))
    
    Application.ScreenUpdating = False
    Call WriteText(3, 6, "")
    Call WriteText(3, 7, "")
    Call WriteText(2, 7, "")
    Call WriteText(2, 8, "")
    Call WriteText(1, 8, "")
    Call WriteText(5, 6, "")
    Call WriteText(5, 7, "")
    Call WriteText(6, 7, "")
    Call WriteText(6, 8, "")
    Call WriteText(7, 8, "")
    Call WriteText(4, 5, "")
    Call WriteText(5, 8, "")
    Application.ScreenUpdating = True
    Application.Wait (Now + TimeValue(wait2))

End Sub

Private Sub Krapf(wait1 As String, wait2 As String, t As String)

    Application.ScreenUpdating = False
    Call WriteText(5, 49, "/")
    Call WriteText(4, 50, "/")
    Call WriteText(3, 51, "/")
    Call WriteText(2, 48, t)
    Application.ScreenUpdating = True
    Application.Wait (Now + TimeValue(wait1))
    
    Application.ScreenUpdating = False
    Call WriteText(5, 49, "")
    Call WriteText(4, 50, "")
    Call WriteText(3, 51, "")
    Call WriteText(2, 48, "")
    Application.ScreenUpdating = True

End Sub

Private Sub Leuerer(wait1 As String, wait2 As String, t As String)

    Application.ScreenUpdating = False
    Call WriteText(4, 28, "\")
    Call WriteText(3, 27, "\")
    Call WriteText(2, 22, t)
    Application.ScreenUpdating = True
    Application.Wait (Now + TimeValue(wait1))
    
    Application.ScreenUpdating = False
    Call WriteText(4, 28, "")
    Call WriteText(3, 27, "")
    Call WriteText(2, 22, "")
    Application.ScreenUpdating = True
    Application.Wait (Now + TimeValue(wait2))

End Sub

Private Sub Opening(wait1 As String, wait2 As String)
    Application.ScreenUpdating = False
    Call WriteText(2, 38, txt(398))
    Call WriteText(3, 38, txt(399))
    Application.ScreenUpdating = True
    Application.Wait (Now + TimeValue(wait1))
    
    Application.ScreenUpdating = False
    Call WriteText(2, 38, "")
    Call WriteText(3, 38, "")
    Application.ScreenUpdating = True
    Application.Wait (Now + TimeValue(wait2))
End Sub

Private Sub BetSlip1()
    Application.ScreenUpdating = False
    Call WriteText(10, 53, "WIN")
    Call WriteText(11, 54, "#1")
    Application.ScreenUpdating = True
End Sub

Private Sub BetSlip2()
    Application.ScreenUpdating = False
    Call WriteText(10, 53, "")
    Call WriteText(11, 54, "")
    Application.ScreenUpdating = True
End Sub

Private Sub Wait(w As String)
    Application.Wait (Now + TimeValue(w))
End Sub

