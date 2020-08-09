VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionsExcel 
   Caption         =   "[Excel options]"
   ClientHeight    =   4380
   ClientLeft      =   72
   ClientTop       =   276
   ClientWidth     =   6084
   OleObjectBlob   =   "frmOptionsExcel.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmOptionsExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Pop-up for setting Excel options
'   UserForm frmOptionsExcel

Private Sub UserForm_Initialize()
    With Me
        'Captions
        .caption = GetText(g_arr_Text, "USERFORM002")
        .Height = 285
        .width = 328
        .chkRS_eo01.caption = GetText(g_arr_Text, "EXCELOPT003")
        .chkRS_eo02.caption = GetText(g_arr_Text, "EXCELOPT004")
        .chkRS_eo03.caption = GetText(g_arr_Text, "EXCELOPT005")
        .chkRS_eo04.caption = GetText(g_arr_Text, "EXCELOPT001")
        .chkRS_eo05.caption = GetText(g_arr_Text, "EXCELOPT002")
        .chkRS_eo06.caption = GetText(g_arr_Text, "EXCELOPT006")
        .chkRS_eo07.caption = GetText(g_arr_Text, "EXCELOPT007")
        .chkRS_eo08.caption = GetText(g_arr_Text, "EXCELOPT008")
        .chkRS_eo09.caption = GetText(g_arr_Text, "EXCELOPT009")
        .cmdRS_eo01.caption = GetText(g_arr_Text, "BTN014")
        .cmdRS_eo02.caption = GetText(g_arr_Text, "BTN019")
        .cmdRS_eo03.caption = GetText(g_arr_Text, "BTN018")
    End With
    
    'Set checkbox values
    Call SetCheckboxes
    
    'Display the UserForm in the center of the Window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
End Sub

Private Sub SetCheckboxes() 'Set the checkbox values
    With Me
        .chkRS_eo01.Value = Application.ActiveWindow.WindowState = xlMaximized 'Checked if maximized, unchecked if normal or minimized
        .chkRS_eo02.Value = Not Application.ActiveWindow.DisplayHeadings
        .chkRS_eo03.Value = Not Application.ActiveWindow.DisplayGridlines
        .chkRS_eo04.Value = Application.DisplayFullScreen
        .chkRS_eo05.Value = Not Application.DisplayFormulaBar
        .chkRS_eo06.Value = Not Application.DisplayStatusBar
        .chkRS_eo07.Value = Not Application.ActiveWindow.DisplayVerticalScrollBar
        .chkRS_eo08.Value = Not Application.ActiveWindow.DisplayHorizontalScrollBar
        .chkRS_eo09.Value = Not Application.ActiveWindow.DisplayWorkbookTabs
    End With
End Sub

Private Sub cmdRS_eo01_Click() 'OK button
    Unload Me
End Sub

Private Sub cmdRS_eo02_Click() 'Reset button
    Call basMainCode.ResetExcelOptions
    Call SetCheckboxes
End Sub

Private Sub cmdRS_eo03_Click() 'Big pic button
    Call basMainCode.ExcelOptionsTVfull
    Call SetCheckboxes
End Sub


'Set Excel options
'-----------------
Private Sub chkRS_eo01_Click() 'Window size
    On Error Resume Next
    
    With Application.ActiveWindow
        If Me.chkRS_eo01 = True Then
            .WindowState = xlMaximized
        Else
            .WindowState = xlNormal
        End If
    End With
End Sub

Private Sub chkRS_eo02_Click() 'Excel row and column headings
    Application.ActiveWindow.DisplayHeadings = Not Me.chkRS_eo02.Value
End Sub

Private Sub chkRS_eo03_Click() 'Excel gridlines
    Application.ActiveWindow.DisplayGridlines = Not Me.chkRS_eo03.Value
End Sub

Private Sub chkRS_eo04_Click() 'Excel ribbon
    Application.DisplayFullScreen = Me.chkRS_eo04.Value
End Sub

Private Sub chkRS_eo05_Click() 'Excel formula bar
    Application.DisplayFormulaBar = Not Me.chkRS_eo05.Value
End Sub

Private Sub chkRS_eo06_Click() 'Excel status bar
    Application.DisplayStatusBar = Not Me.chkRS_eo06.Value
End Sub

Private Sub chkRS_eo07_Click() 'Excel vertical scrollbar
    Application.ActiveWindow.DisplayVerticalScrollBar = Not Me.chkRS_eo07.Value
End Sub

Private Sub chkRS_eo08_Click() 'Excel horizontal scrollbar
    Application.ActiveWindow.DisplayHorizontalScrollBar = Not Me.chkRS_eo08.Value
End Sub

Private Sub chkRS_eo09_Click() 'Excel workbook tabs
    Application.ActiveWindow.DisplayWorkbookTabs = Not Me.chkRS_eo09.Value
End Sub
