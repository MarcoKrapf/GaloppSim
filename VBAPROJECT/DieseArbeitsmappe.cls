VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    
    'Set g_strPlayMode according to the file type
        If ThisWorkbook.FileFormat = 55 Then '.xlam
            g_strPlayMode = "AI"
        Else 'FileFormat = 52 (.xlsm)
            g_strPlayMode = "RS"
        End If
        
    'Set the standard settings for the race options
        g_blnStopAlgorithms = False
        g_blnBetsPlaced = False
        g_blnTactics0 = False
        g_blnTactics3 = True
        g_blnTactics6 = False
        g_blnIncidentRefuse = True
        g_blnSlipstream = False
        g_blnSlipstreamDouble = True
        g_blnSlipstreamShow = True
        g_blnFocusedRun = False
        g_blnHighlightFoc = True
        g_blnHoofprints = False
        g_blnHorseNamesLeft = True
        g_blnHorseColoursLeft = True
        g_blnHighlightFav = True
        g_blnHorseNamesFinish = True
        g_blnRankingColours = True
        g_blnRankingDelay = False
        g_blnBettingMode = True
        g_blnBettingAnalysis = True
        g_intTrackMetres = 250
        g_blnRaceInformation = True
        g_blnRaceInfoPopup = False
        g_blnRaceInfoWorksheet = True
        g_blnRaceInfoLeader = True
        g_blnRaceInfoProgressBar = True
        g_lngRaceInfoBackColour = vbWhite
        g_lngRaceInfoForeColour = vbBlack
        g_intGMPspeedfactor = 1
        g_strExcelMode = "normal"
        g_blnSpeech = False
        
    'Set the default path for GaloppSim files
        g_defaultPath = Application.UserLibraryPath
   
    'Determine the standard language according to the Excel version
        Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
            Case msoLanguageIDGerman, msoLanguageIDGermanAustria, _
                    msoLanguageIDGermanLiechtenstein, msoLanguageIDGermanLuxembourg
                g_strLanguage = "DE"
            Case msoLanguageIDBulgarian
                g_strLanguage = "BG"
            Case Else
                g_strLanguage = "EN"
        End Select
    
    If g_strPlayMode = "RS" Then
        Call NewRace_RS
    End If
    
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
'    'Reset Excel options
'        Call basMainCode.ResetExcelOptions
'        Application.ScreenUpdating = True
'    If g_strPlayMode = "RS" Then
'        'Do not save the workbook
'        'https://support.microsoft.com/de-de/help/213428/how-to-suppress-save-changes-prompt-when-you-close-a-workbook-in-excel
'        ThisWorkbook.Saved = True
'    End If
End Sub


