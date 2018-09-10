Attribute VB_Name = "basMovie2017"
Option Explicit
Option Private Module

'This animation shows the original scene in the afternoon of 2 July 2017,
'when the idea for programming the Excel horse racing simulator was born

Public Sub PlayMovie2017(ActSheet As String)
    
    Dim m_wksCheck As Worksheet
    Dim strL As String 'Language of the picture texts (German or English)
    
    'Determine the language
    If g_strLanguage = "DE" Or g_strLanguage = "CH" Then
        strL = "DE"
    Else
        strL = "EN"
    End If
    
    'Check whether the worksheet "GALOPPSIM_MOVIE" already exists
    For Each m_wksCheck In ActiveWorkbook.Worksheets
        If m_wksCheck.name = "GALOPPSIM_MOVIE" Then 'Worksheet exists
            Application.DisplayAlerts = False
            m_wksCheck.Delete 'Delete the worksheet
            Application.DisplayAlerts = True
        End If
    Next m_wksCheck
    
    'Create the worksheet "GALOPPSIM_MOVIE"
        Set g_wksMovie = ActiveWorkbook.Worksheets.Add(Before:=Sheets(1))
        With g_wksMovie
            .name = "GALOPPSIM_MOVIE"
            .Range(Columns(1), Columns(100)).ColumnWidth = 2
            .Activate
        End With
        
        If g_strPlayMode = "AI" Then Call basAuxiliary.AI_ExcelModeStart
        
    'Show a pop-up if the window size is too small for the movie
        If Application.ActiveWindow.Height < 780 Or Application.ActiveWindow.Width < 1080 _
            And g_strExcelMode <> "TVfull" Then
                'Set the button mode
                g_strMsgButtons = "YesNo"
                'Assign the text for the pop-up
                g_strMsgCaption = GetTxt(g_arrTxt, "USERFORM004")
                g_strMsgText = GetTxt(g_arrTxt, "MOVIE001") & vbNewLine & GetTxt(g_arrTxt, "MOVIE002")
                'Display the pop-up
                frmMsg_MultiPurpose.Show
                'Evaluate the return value
                If g_strButtonPressed = "YES" Then Call ExcelOptionsTVfull 'Activate the full screen mode
        End If
    
    'Prepare the speaker
    g_wksMovie.Cells(4, 5).Font.name = "MV Boli"
    g_wksMovie.Cells(5, 8).Font.name = "MV Boli"
    g_wksMovie.Range(Cells(2, 38), Cells(3, 38)).Font.name = "Arial Rounded MT Bold"
    g_wksMovie.Range(Cells(2, 28), Cells(5, 51)).Font.name = "Arial Rounded MT Bold"
    g_wksMovie.Range(Cells(2, 22), Cells(4, 28)).Font.name = "Arial Rounded MT Bold"
    With g_wksMovie.Cells(20, 59)
        .Value = GetTxt(g_arrTxt, "MOVIE005")
    End With
    
    'Play the title sequence
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A0BLACK"))
    Call Wait("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A1" & strL))
    Call Wait("0:00:03")
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A0BLACK"))
    Call Wait("0:00:01")

    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A2" & strL))
    Call Wait("0:00:03")
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A0BLACK"))
    Call Wait("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A3" & strL))
    Call Wait("0:00:03")
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A0BLACK"))
    Call Wait("0:00:01")
    
    'Play the movie
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_01"))
    Call Wait("0:00:02")
    Call Opening("0:00:04", "0:00:02")
    Call Speaker("0:00:02", "0:00:01", GetTxt(g_arrTxt, "MOVIE006"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_02"))
    Call ShowBetSlip
    Call SpeakMatjes("0:00:02", "0:00:02", GetTxt(g_arrTxt, "MOVIE007"))
    If g_blnSpeech Then Call SpeechOut(GetTxt(g_arrTxt, "MOVIE007"))
    Call HideBetSlip
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_01"))
    Call Wait("0:00:02")
    Call Speaker("0:00:02", "0:00:01", GetTxt(g_arrTxt, "MOVIE008"))
    Call Speaker("0:00:02", "0:00:01", GetTxt(g_arrTxt, "MOVIE009"), GetTxt(g_arrTxt, "MOVIE010"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE011"), GetTxt(g_arrTxt, "MOVIE012"))
'    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE013"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE014"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_03"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE015"), GetTxt(g_arrTxt, "MOVIE016"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_04"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE017"), GetTxt(g_arrTxt, "MOVIE018"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_05"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE019"), GetTxt(g_arrTxt, "MOVIE020"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_06"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE021"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_07"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE022"), GetTxt(g_arrTxt, "MOVIE023"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_01"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE024"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE025"), GetTxt(g_arrTxt, "MOVIE026"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE027"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE028"), GetTxt(g_arrTxt, "MOVIE029"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE030"), GetTxt(g_arrTxt, "MOVIE031"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE032"), GetTxt(g_arrTxt, "MOVIE033"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_08"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE034"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_09"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE035"), GetTxt(g_arrTxt, "MOVIE036"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_10"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE037"), GetTxt(g_arrTxt, "MOVIE038"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_11"))
    g_wksMovie.Range(Cells(1, 5), Cells(7, 8)).Font.FontStyle = "Bold"
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE039"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_12"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE040"), GetTxt(g_arrTxt, "MOVIE041"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_13"))
    Call Speaker("0:00:02", "0:00:00", GetTxt(g_arrTxt, "MOVIE042"), GetTxt(g_arrTxt, "MOVIE043"))
    g_wksMovie.Range(Cells(1, 5), Cells(7, 8)).Font.FontStyle = "Regular"
    Call Speaker("0:00:02", "0:00:02", GetTxt(g_arrTxt, "MOVIE044"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_14"))
    Call Wait("0:00:02")
    Call SpeakMatjes("0:00:04", "0:00:02", GetTxt(g_arrTxt, "MOVIE045"))
    If g_blnSpeech Then Call SpeechOut(GetTxt(g_arrTxt, "MOVIE045"))
    Call SpeakFlo("0:00:04", "0:00:02", GetTxt(g_arrTxt, "MOVIE046"))
    If g_blnSpeech Then Call SpeechOut(GetTxt(g_arrTxt, "MOVIE046"))
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_15"))
    Call Wait("0:00:02")
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_16"))
    Call Wait("0:00:02")

    'Play the closing credits
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A0BLACK"))
    Call Wait("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_Z1" & strL))
    Call Wait("0:00:03")
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A0BLACK"))
    Call Wait("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_Z3" & strL))
    Call Wait("0:00:03")
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A0BLACK"))
    Call Wait("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_Z4" & strL))
    Call Wait("0:00:03")
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A0BLACK"))
    Call Wait("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_Z7" & strL))
    Call Wait("0:00:03")
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A0BLACK"))
    Call Wait("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_Z8" & strL))
    Call Wait("0:00:03")
    Call DrawPicture(basAuxiliary.GetPictureColumn("MOVIE1_A0BLACK"))
    Call Wait("0:00:01")
    
    'Delete the movie worksheet
        Application.DisplayAlerts = False 'Suppress the warning message
        ActiveWorkbook.Worksheets("GALOPPSIM_MOVIE").Delete 'Delete the movie worksheet
        Application.DisplayAlerts = True 'Activate warning messages

    'If the full screen mode was activated: reset the Excel options
        If g_strButtonPressed = "YES" Then Call ResetExcelOptions
        
    'Reset AI Excel mode
        If g_strPlayMode = "AI" Then Call basAuxiliary.AI_ExcelModeEnd
    
    'Jump to the main worksheet
        If g_strPlayMode = "RS" Then g_wksRace.Activate

End Sub

Private Sub DrawPicture(col As Integer)

    Dim i As Integer, j As Integer
    Dim row As Integer

    row = 2 'set the initial row for reading the picture
    
    Application.ScreenUpdating = False 'Deactivate screen updating
    
    For i = 1 To 40
        For j = 1 To 100
            g_wksMovie.Cells(i, j).Interior.Color = g_wksPicData.Cells(row, col).Value
            row = row + 1 'next row on the worksheet "Pic"
        Next
    Next
    
    g_wksMovie.Cells(41, 1).Select
    Application.ScreenUpdating = True 'Activate screen updating

End Sub

Private Sub WriteText(r As Integer, c As Integer, text As String)
    g_wksMovie.Cells(r, c).Value = text
End Sub

Private Sub Speaker(waitBefore As String, waitAfter As String, text1 As String, Optional text2 As String)

    Application.ScreenUpdating = False 'Deactivate screen updating
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
    Call WriteText(4, 5, text1)
    Call WriteText(5, 8, text2)
    Application.ScreenUpdating = True 'Activate screen updating
    Application.Wait (Now + TimeValue(waitBefore))
    
    Application.ScreenUpdating = False 'Deactivate screen updating
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
    Application.ScreenUpdating = True 'Activate screen updating
    Application.Wait (Now + TimeValue(waitAfter))

End Sub

Private Sub SpeakMatjes(wait1 As String, wait2 As String, text As String)

    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(5, 49, "/")
    Call WriteText(4, 50, "/")
    Call WriteText(3, 51, "/")
    Call WriteText(2, 46, text)
    Application.ScreenUpdating = True 'Activate screen updating
    Application.Wait (Now + TimeValue(wait1))
    
    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(5, 49, "")
    Call WriteText(4, 50, "")
    Call WriteText(3, 51, "")
    Call WriteText(2, 46, "")
    Application.ScreenUpdating = True 'Activate screen updating

End Sub

Private Sub SpeakFlo(wait1 As String, wait2 As String, text As String)

    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(4, 28, "\")
    Call WriteText(3, 27, "\")
    Call WriteText(2, 22, text)
    Application.ScreenUpdating = True 'Activate screen updating
    Application.Wait (Now + TimeValue(wait1))
    
    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(4, 28, "")
    Call WriteText(3, 27, "")
    Call WriteText(2, 22, "")
    Application.ScreenUpdating = True 'Activate screen updating
    Application.Wait (Now + TimeValue(wait2))

End Sub

Private Sub Opening(waitBefore As String, waitAfter As String)
    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(2, 38, GetTxt(g_arrTxt, "MOVIE003"))
    Call WriteText(3, 38, GetTxt(g_arrTxt, "MOVIE004"))
    Application.ScreenUpdating = True 'Activate screen updating
    Application.Wait (Now + TimeValue(waitBefore))
    
    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(2, 38, "")
    Call WriteText(3, 38, "")
    Application.ScreenUpdating = True 'Activate screen updating
    Application.Wait (Now + TimeValue(waitAfter))
End Sub

Private Sub ShowBetSlip()
    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(10, 53, "WIN")
    Call WriteText(11, 54, "#1")
    Application.ScreenUpdating = True 'Activate screen updating
End Sub

Private Sub HideBetSlip()
    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(10, 53, "")
    Call WriteText(11, 54, "")
    Application.ScreenUpdating = True 'Activate screen updating
End Sub

Private Sub Wait(waitSeconds As String) 'Procedure for delay
    Application.Wait (Now + TimeValue(waitSeconds))
End Sub
