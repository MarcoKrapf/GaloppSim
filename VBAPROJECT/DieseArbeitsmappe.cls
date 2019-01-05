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
    
    'Set PlayMode according to the file type
        If ThisWorkbook.FileFormat = 55 Then '.xlam
            g_strPlayMode = "AI"
        Else 'FileFormat = 52 (.xlsm)
            g_strPlayMode = "RS"
        End If
        
    'Create objects
        Set objRace = New clsRace 'Object for the race data
        Set objOption = New clsOptions 'Object for the race and excel optons
        
    'Set the standard settings for the race options
        objOption.STOP_ALG = False
        objOption.EXCEL_MODE = "normal"
        objOption.TACTICS = 3
        objOption.REFUSE_RUN = True
        objOption.SLIPSTREAM = True
        objOption.SLIPSTREAM_DBL = False
        objOption.SLIPSTREAM_SHOW = False
        objOption.FOCUSED_RUN = False
        objOption.HIGHLIGHT_FOC = True
        objOption.HOOFPRINTS = False
        objOption.NAMES_LEFT = True
        objOption.COLOURS_LEFT = True
        objOption.HIGHLIGHT_FAV = True
        objOption.NAMES_FINISH = True
        objOption.NAMES_PHOTO = True
        objOption.PHOTO_BW = False
        objOption.RANKING_COL = True
        objOption.RANKING_DELAY = False
        objOption.BET_MODE = True
        objOption.BET_PLACED = False
        objOption.BET_ANALYSIS = True
        objOption.METRES_DISPLAY = 250
        objOption.RACE_INFO = True
        objOption.RACE_INFO_POP = False
        objOption.RACE_INFO_WKS = True
        objOption.RACE_INFO_LEADER = True
        objOption.RACE_INFO_PROGRESS = True
        objOption.RACE_INFO_COL_B = vbWhite
        objOption.RACE_INFO_COL_F = vbBlack
        objOption.SPEED_FACTOR = 1
        objOption.SPEECH = False
        objOption.TIDE = 1
        objOption.LUGWORMS = 1
        
    'Set the default path for GaloppSim files
        g_defaultPath = Application.UserLibraryPath
   
    'Determine the standard language according to the Excel version
        Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
            'German speaking countries
            Case msoLanguageIDGerman, msoLanguageIDGermanAustria, _
                    msoLanguageIDGermanLiechtenstein, msoLanguageIDGermanLuxembourg
                objOption.language = "DE"
            'Bulgaria
            Case msoLanguageIDBulgarian
                objOption.language = "BG"
            'Russian speaking countries
'            Case msoLanguageIDRussian, msoLanguageIDUkrainian
'                objOption.LANGUAGE = "RU"
            'All other
            Case Else
                objOption.language = "EN"
        End Select
    
    If g_strPlayMode = "RS" Then
        Call RS_NewRace
    End If
    
End Sub

'Execute this procedure when the workbook is being closed.
'The Workbook_BeforeClose event can be used alternatively to the Auto_Close
'procedure which is used for this project in the module "basMainCode".
'If both procedures are implemented first the Workbook_BeforeClose is executed
'followed by Auto_Close.
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

