Attribute VB_Name = "basMovie2017"
Option Explicit
Option Private Module

'This module contains an animation showing the original scene in the afternoon
'of 2 July 2017, when the idea for programming the Excel Horse Racing Simulator was born
'   Module basMovie2017

Public Sub PlayMovie2017()
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    Dim m_wksCheck As Worksheet
    Dim strL As String 'Language of the texts
    
    'Determine the language (German or English)
    If objOption.language = "DE" Or objOption.language = "CH" Then 'German speaking countries
        strL = "DE"
    Else 'All other
        strL = "EN"
    End If
    
    'Check whether the Worksheet "GALOPPSIM_MOVIE" already exists
    For Each m_wksCheck In ActiveWorkbook.Worksheets
        If m_wksCheck.name = "GALOPPSIM_MOVIE" Then 'If the Worksheet exists
            Application.DisplayAlerts = False
            m_wksCheck.Delete 'Delete the Worksheet
            Application.DisplayAlerts = True
        End If
    Next m_wksCheck
    
    'Create the Worksheet "GALOPPSIM_MOVIE"
        Set g_wksMovie = ActiveWorkbook.Worksheets.Add(Before:=Sheets(1))
        With g_wksMovie
            .name = "GALOPPSIM_MOVIE"
            .Range(Columns(1), Columns(100)).ColumnWidth = 2
            .Activate
        End With
        
        If g_strPlayMode = "AI" Then Call AI_ExcelModeStart
        
    'Show a pop-up if the window size is too small for the movie
        If (Application.ActiveWindow.Height < 780 Or Application.ActiveWindow.width < 1080) _
            And objOption.EXCEL_MODE <> "TVfull" Then
            
            Call ShowMessagePopup(GetText(g_arr_Text, "USERFORM004"), _
                GetText(g_arr_Text, "MOVIE001") & vbNewLine _
                    & GetText(g_arr_Text, "MOVIE002"), enumButton.YesNo, vbModal)
                
            'Evaluate the return value of the button
            If g_enumButton = enumButton.yes Then Call ExcelOptionsTVfull 'Activate the full screen mode
        End If
    
    'Prepare the speaker
    With g_wksMovie
        .Cells(4, 5).Font.name = "MV Boli"
        .Cells(5, 8).Font.name = "MV Boli"
        .Range(Cells(2, 38), Cells(3, 38)).Font.name = "Arial Rounded MT Bold"
        .Range(Cells(2, 28), Cells(5, 51)).Font.name = "Arial Rounded MT Bold"
        .Range(Cells(2, 22), Cells(4, 28)).Font.name = "Arial Rounded MT Bold"
        .Cells(20, 59).Value = GetText(g_arr_Text, "MOVIE005")
    End With
    
    'Play the title sequence
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A0BLACK"))
    Call Delay("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A1" & strL))
    Call Delay("0:00:03")
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A0BLACK"))
    Call Delay("0:00:01")

    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A2" & strL))
    Call Delay("0:00:03")
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A0BLACK"))
    Call Delay("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A3" & strL))
    Call Delay("0:00:03")
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A0BLACK"))
    Call Delay("0:00:01")
    
    'Play the movie
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_01"))
    Call Delay("0:00:02")
    Call Opening("0:00:04", "0:00:02")
    Call Speaker("0:00:02", "0:00:01", GetText(g_arr_Text, "MOVIE006"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_02"))
    Call ShowBetSlip
    Call SpeakMatjes("0:00:02", GetText(g_arr_Text, "MOVIE007"))
    If objOption.SPEECH Then Call SpeechOut(GetText(g_arr_Text, "MOVIE007"))
    Call HideBetSlip
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_01"))
    Call Delay("0:00:02")
    Call Speaker("0:00:02", "0:00:01", GetText(g_arr_Text, "MOVIE008"))
    Call Speaker("0:00:02", "0:00:01", GetText(g_arr_Text, "MOVIE009"), GetText(g_arr_Text, "MOVIE010"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE011"), GetText(g_arr_Text, "MOVIE012"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE014"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_03"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE015"), GetText(g_arr_Text, "MOVIE016"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_04"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE017"), GetText(g_arr_Text, "MOVIE018"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_05"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE019"), GetText(g_arr_Text, "MOVIE020"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_06"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE021"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_07"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE022"), GetText(g_arr_Text, "MOVIE023"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_01"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE024"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE025"), GetText(g_arr_Text, "MOVIE026"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE027"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE028"), GetText(g_arr_Text, "MOVIE029"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE030"), GetText(g_arr_Text, "MOVIE031"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE032"), GetText(g_arr_Text, "MOVIE033"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_08"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE034"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_09"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE035"), GetText(g_arr_Text, "MOVIE036"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_10"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE037"), GetText(g_arr_Text, "MOVIE038"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_11"))
    g_wksMovie.Range(Cells(1, 5), Cells(7, 8)).Font.FontStyle = "Bold"
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE039"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_12"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE040"), GetText(g_arr_Text, "MOVIE041"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_13"))
    Call Speaker("0:00:02", "0:00:00", GetText(g_arr_Text, "MOVIE042"), GetText(g_arr_Text, "MOVIE043"))
    g_wksMovie.Range(Cells(1, 5), Cells(7, 8)).Font.FontStyle = "Regular"
    Call Speaker("0:00:02", "0:00:02", GetText(g_arr_Text, "MOVIE044"))
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_14"))
    Call Delay("0:00:02")
    Call SpeakMatjes("0:00:04", GetText(g_arr_Text, "MOVIE045"))
    If objOption.SPEECH Then Call SpeechOut(GetText(g_arr_Text, "MOVIE045"))
    Call SpeakFlo("0:00:04", GetText(g_arr_Text, "MOVIE046"))
    If objOption.SPEECH Then Call SpeechOut(GetText(g_arr_Text, "MOVIE046"))
    Call Delay("0:00:02")
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_15"))
    Call Delay("0:00:02")
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_16"))
    Call Delay("0:00:02")

    'Play the closing credits
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A0BLACK"))
    Call Delay("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_Z1" & strL))
    Call Delay("0:00:03")
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A0BLACK"))
    Call Delay("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_Z3" & strL))
    Call Delay("0:00:03")
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A0BLACK"))
    Call Delay("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_Z4" & strL))
    Call Delay("0:00:03")
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A0BLACK"))
    Call Delay("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_Z7" & strL))
    Call Delay("0:00:03")
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A0BLACK"))
    Call Delay("0:00:01")
    
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_Z8" & strL))
    Call Delay("0:00:03")
    Call DrawPicture(basAuxiliary.GetColumn(g_wksPIC, "MOVIE1_A0BLACK"))
    Call Delay("0:00:01")
    
    'Delete the Worksheet
        Application.DisplayAlerts = False 'Suppress the warning message
        ActiveWorkbook.Worksheets("GALOPPSIM_MOVIE").Delete 'Delete the Worksheet
        Application.DisplayAlerts = True 'Re-activate warning messages

    On Error Resume Next
    'If the full screen mode was activated: Reset the Excel options
        If g_enumButton = enumButton.yes Then Call ResetExcelOptions
        
    'Reset AI Excel mode
        If g_strPlayMode = "AI" Then Call AI_ExcelModeEnd
    
    'Jump to the GALOPPSIM Worksheet
        If g_strPlayMode = "RS" Then g_wksRace.Activate
        
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "PlayMovie2017()")
    Call basAuxiliary.CodeCrash
End Sub

Private Sub DrawPicture(col As Integer)

    Dim i As Integer, j As Integer
    Dim row As Integer

    row = 2 'Initial row for reading the colour data for the picture
    
    Application.ScreenUpdating = False 'Deactivate screen updating
    
    For i = 1 To 40
        For j = 1 To 100
            g_wksMovie.Cells(i, j).Interior.color = g_wksPIC.Cells(row, col).Value
            row = row + 1 'Next row on the worksheet "PIC"
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

Private Sub SpeakMatjes(waitTime As String, text As String)

    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(5, 49, "/")
    Call WriteText(4, 50, "/")
    Call WriteText(3, 51, "/")
    Call WriteText(2, 46, text)
    Application.ScreenUpdating = True 'Activate screen updating
    Application.Wait (Now + TimeValue(waitTime))
    
    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(5, 49, "")
    Call WriteText(4, 50, "")
    Call WriteText(3, 51, "")
    Call WriteText(2, 46, "")
    Application.ScreenUpdating = True 'Activate screen updating

End Sub

Private Sub SpeakFlo(waitTime As String, text As String)

    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(4, 28, "\")
    Call WriteText(3, 27, "\")
    Call WriteText(2, 22, text)
    Application.ScreenUpdating = True 'Activate screen updating
    Application.Wait (Now + TimeValue(waitTime))
    
    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(4, 28, "")
    Call WriteText(3, 27, "")
    Call WriteText(2, 22, "")
    Application.ScreenUpdating = True 'Activate screen updating

End Sub

Private Sub Opening(waitBefore As String, waitAfter As String)

    Application.ScreenUpdating = False 'Deactivate screen updating
    Call WriteText(2, 38, GetText(g_arr_Text, "MOVIE003"))
    Call WriteText(3, 38, GetText(g_arr_Text, "MOVIE004"))
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

'Procedure for delay
Private Sub Delay(waitSeconds As String)
    Application.Wait (Now + TimeValue(waitSeconds))
End Sub
