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
        
    'Set the standard settings for the racing options
        g_blnBetsPlaced = False
        g_blnTactics0 = False
        g_blnTactics3 = True
        g_blnTactics6 = False
        g_blnFocusedRun = False
        g_blnHighlightFoc = True
        g_blnHoofprints = False
        g_blnHorseNamesLeft = True
        g_blnHorseColoursLeft = True
        g_blnHighlightFav = True
        g_blnHorseNamesFinish = True
        g_blnRankingColours = True
        g_blnRankingDelay = False
        g_blnRaceInformation = True
        g_blnBettingMode = True
        g_blnBettingAnalysis = True
   
    'Determine the standard language according to the Excel version
        Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
            Case msoLanguageIDGerman, msoLanguageIDGermanAustria, _
                    msoLanguageIDGermanLiechtenstein, msoLanguageIDGermanLuxembourg
                g_strLanguage = "DE"
'            Case msoLanguageIDRussian, msoLanguageIDUkrainian
'                g_strLanguage = "RU"
            Case Else
                g_strLanguage = "EN"
        End Select

    'No m_strRaceName chosen
        g_strRaceSelected = ""
    
    If g_strPlayMode = "RS" Then Call NewRace_RS
    
End Sub
