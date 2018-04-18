VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionsExcel 
   Caption         =   "[Excel options]"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6360
   OleObjectBlob   =   "frmOptionsExcel.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmOptionsExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    With Me
        'Captions
        .Caption = g_strTxt(253)
        .chkRS_eo01.Caption = g_strTxt(271)
        .chkRS_eo02.Caption = g_strTxt(272)
        .chkRS_eo03.Caption = g_strTxt(273)
        .chkRS_eo04.Caption = g_strTxt(269)
        .chkRS_eo05.Caption = g_strTxt(270)
        .chkRS_eo06.Caption = g_strTxt(274)
        .chkRS_eo07.Caption = g_strTxt(275)
        .chkRS_eo08.Caption = g_strTxt(276)
        .chkRS_eo09.Caption = g_strTxt(277)
        .cmdRS_eo01.Caption = g_strTxt(180)
        .cmdRS_eo02.Caption = g_strTxt(185)
        .cmdRS_eo03.Caption = g_strTxt(184)
        'ControlTipTexts
        .chkRS_eo01.ControlTipText = g_strTxt(319)
        .chkRS_eo02.ControlTipText = g_strTxt(321)
        .chkRS_eo03.ControlTipText = g_strTxt(323)
        .chkRS_eo04.ControlTipText = g_strTxt(345)
        .chkRS_eo05.ControlTipText = g_strTxt(347)
        .chkRS_eo06.ControlTipText = g_strTxt(349)
        .chkRS_eo07.ControlTipText = g_strTxt(351)
        .chkRS_eo08.ControlTipText = g_strTxt(353)
        .chkRS_eo09.ControlTipText = g_strTxt(355)
    End With
    
    'Set checkbox values
    Call SetCheckboxes
    
    'Hide the ribbon checkbox in AI edition
    If g_strPlayMode = "AI" Then chkRS_eo04.Visible = False
    
    'Display the UserForm in the center of the Window
    Call basMainCode.PlaceUserFormInCenter(Me)
End Sub

Private Sub SetCheckboxes() 'Set checkbox values
    With Me
        .chkRS_eo01.Value = Application.ActiveWindow.WindowState = xlMaximized 'checked if maximized, unchecked if normal or minimized
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
    Call basMainCode.BigPicExcelOptions
    Call SetCheckboxes
End Sub


'Set Excel options
'-----------------
Private Sub chkRS_eo01_Click() 'Window size
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
