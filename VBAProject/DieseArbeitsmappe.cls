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

'Workbook events
'   ThisWorkbook (DieseArbeitsmappe)

'This procedure is executed on opening the Excel Workbook
Private Sub Workbook_Open()
    'As this procedure is triggered automatically
    'when opening the workbook comment in the "Stop" command in the
    'following line for debugging purposes
'    Stop
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    'Values for development purposes - save and re-open the Workbook for taking effect
    'or change the settings in the module basDeveloper>DevelopmentTools
        g_skipDelay = False 'Set true to skip delay commands (Application.Wait)
        g_errorLogPath = Environ("UserProfile") '& "\Desktop" 'Path for the error log file
        g_errorLogging = False 'Error logging on(true)/off(false)
        g_payoutLogging = False 'Pay-out logging for bets on/off
    'If error logging is turned on: Activate a(ny) VBComponent to have a pane object.
    'This is needed to get the CodeModule name in the error handlers.
        If g_errorLogging = True Then
            On Error Resume Next 'Skip errors
            'Check trust access to the VBA project object model
            If TrustAccessToVBAprojectObjectModel = True Then 'Trust access
                Application.VBE.ActiveVBProject.VBComponents("DUMMY").Activate 'Take a dummy
            Else 'No trust access
                MsgBox "No trust access to the VBA project object model." & vbNewLine _
                    & "Error logging is deactivated."
                    g_errorLogging = False
            End If
            On Error GoTo ERRORHANDLING 'Stop skipping errors
        End If
    
    'Set the PlayMode according to the file type
        If ThisWorkbook.IsAddin Then
            g_strPlayMode = "AI"
        Else
            g_strPlayMode = "RS"
        End If
'        'Alternatively: Check the file format (.xlam = 55 // .xlsm = 52)
'        If ThisWorkbook.FileFormat = 55 Then
'            g_strPlayMode = "AI"
'        Else
'            g_strPlayMode = "RS"
'        End If
        
    'Assign the worksheets of this workbook
        Set g_wksTEXT = VBAProject_GaloppSim.Table_TEXT  'Take the internal name (CodeName)
        Set g_wksPIC = VBAProject_GaloppSim.Table_PIC
        Set g_wksTEC = VBAProject_GaloppSim.Table_TEC
        
'        'Alternative but less stable assignment:
'        Set g_wksTEC = ThisWorkbook.Worksheets("TEC") 'Attenton: The external name (Worksheet name) could be changed by the user

    'Create objects of Class Modules
        Set objRace = New clsRace 'Object for the race data
        Set objOption = New clsOptions 'Object for the race and excel options
        Set objSpeed = New clsSpeed 'Object for speed data
        Set objStat = New clsStatisctics 'Object for statistical data
        
    'Set the standard settings for the race options
        objOption.STOP_ALG = False
        objOption.EXCEL_MODE = "normal"
        objOption.ZOOM_LEVEL = 3
        objOption.TACTICS = True
        objOption.TACTICS_REVEAL_TAC = False
        objOption.TACTICS_REVEAL_CURR = False
        objOption.REFUSE_RUN = True
        objOption.REFUSAL_RATE = 100
        objOption.SLIPSTREAM = True
        objOption.SLIPSTREAM_DBL = False
        objOption.SLIPSTREAM_SHOW = False
        objOption.FOCUSED_RUN = enumCamera.standard
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
        objOption.MOMENTUM = False
        objOption.MOMENTUM_REFRESHRATE = 4
        objOption.SPEECH = False
        objOption.TIDE = 1
        objOption.LUGWORMS = 1
        objOption.SPACE_PLANET = enumPlanets.moon
        objOption.SPACE_ALIENS = enumAliens.unfriendly
        objOption.SPACE_KIDNAPPINGRATE = 2
        objOption.TRIBUNES = True
        objOption.SPECTATORS = 0 'No spectators allowed during the Corona crisis
        objOption.PARTICULATES_SLIDER = 3
        objOption.COL_BACK = xlNone
        objOption.COL_TEXT = vbBlack
        objOption.COL_RANKINGS = vbWhite
        objOption.DAYLIGHT = 0 'High noon (12:00)
        
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
    
    'Get texts according to the selected language
    GetTextComponents 'Call the Public Sub 'GetTextComponents' in module 'basAuxiliary'
'    'Alternatively: More illustrative using the 'Call' command
'    Call GetTextComponents
'    'Alternatively: Even more illustrative using the module name and the 'Call' command
'    Call basAuxiliary.GetTextComponents
    
    If g_strPlayMode = "RS" Then
        Application.ActiveWindow.WindowState = xlMaximized 'Maximize the window
        Call RS_NewRace
    End If
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "Workbook_Open()")
    Call basAuxiliary.CodeCrash
End Sub

'This procedure is executed when the Workbook is being closed.
'The Workbook_BeforeClose event can be used alternatively to the Auto_Close
'procedure which is used for this project in the module "basMainCode".
'If both procedures are implemented first the Workbook_BeforeClose is executed
'followed by Auto_Close.
Private Sub Workbook_BeforeClose(Cancel As Boolean)
'    'Reset all Excel options
'    Call basMainCode.ResetExcelOptions
'    Application.ScreenUpdating = True
'    If g_strPlayMode = "RS" Then
'        ThisWorkbook.Saved = True 'Do not save the workbook!
'        'https://support.microsoft.com/en-us/help/213428/how-to-suppress-save-changes-prompt-when-you-close-a-workbook-in-excel
'    End If
End Sub

'Check whether the checkbox "Trust access to the VBA project object model"
'in the "Developer Macro Settings" of the "Trust Center" is ticked
Private Function TrustAccessToVBAprojectObjectModel() As Boolean
    On Error GoTo no_trust
    
    Dim a As Integer
    a = ThisWorkbook.VBProject.VBComponents.count 'Error if no trust access
    
    TrustAccessToVBAprojectObjectModel = True
    
    Exit Function
no_trust:
    TrustAccessToVBAprojectObjectModel = False
End Function
