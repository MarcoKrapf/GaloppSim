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
    
    'Initialize the Super Developer Array
    g_arr_Developer(1) = 0
    g_arr_Developer(2) = 0
    g_arr_Developer(3) = 0
    
    'Values for development purposes - save and re-open the Workbook for taking effect
    'or change the settings in the module basDeveloper>DevelopmentTools
        g_skipDelay = False 'True 'Set true to skip delay commands (Application.Wait)
        g_errorLogPath = Environ("UserProfile") '& "\Desktop" 'Path for the error log file
        g_errorLogging = False 'Error logging on(true)/off(false)
        g_payoutLogging = False 'Pay-out logging for bets on/off
        g_defaultMLpath = Environ("UserProfile") 'Default path for Machine Learning export
        g_MLdataFileName = "GALOPPSIM_ML_DATA" 'File name for Machine Learning export
        g_defaultAutoSavePath = Environ("UserProfile") 'Default path for the auto-save function after a race
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
        Set g_wksTCASE = VBAProject_GaloppSim.Table_TESTCASES
        
'        'Alternative but less stable assignment:
'        Set g_wksTEC = ThisWorkbook.Worksheets("TEC") 'Attenton: The external name (Worksheet name) could be changed by the user

    'Create objects of Class Modules
        Set objBasicData = New clsBasicData 'Object for the basic data
        Set objRace = New clsRace 'Object for the race data
        Set objOption = New clsOptions 'Object for the race and excel options
        Set objSpeed = New clsSpeed 'Object for speed data
        Set objStat = New clsStatisctics 'Object for statistical data
        
    'Set the standard settings for the race options
        With objOption
            .STOP_ALG = False
            .EXCEL_MODE = "normal"
            .TACTICS = True
            .TACTICS_REVEAL_TAC = False
            .TACTICS_REVEAL_CURR = False
            .REFUSE_RUN = True
            .REFUSAL_RATE = 100
            .SLIPSTREAM_IMPACT = 1
            .SLIPSTREAM_SHOW = False
            .FOCUSED_RUN = enumCamera.standard
            .HIGHLIGHT_FOC = True
            .HOOFPRINTS = False
            .NAMES_LEFT = True
            .COLOURS_LEFT = True
            .ANNOUNCE_FAV = True
            .HIGHLIGHT_FAV = True
            .NAMES_FINISH = True
            .NAMES_PHOTO = True
            .PHOTO_BW = False
            .RANKING_COL = True
            .RANKING_DELAY = False
            .BET_MODE = True
            .BET_PLACED = False
            .BET_ANALYSIS = True
            .METRES_DISPLAY = 250
            .RACE_INFO = True
            .RACE_INFO_POP = True
            .RACE_INFO_WKS = False
            .RACE_INFO_LEADER = True
            .RACE_INFO_PROGRESS = True
            .RACE_INFO_COL_B = vbWhite
            .RACE_INFO_COL_F = vbBlack
            .SPEED_FACTOR = 1
            .MOMENTUM_BARS = False
            .MOMENTUM_ICONS = False
            .MOMENTUM_REFRESHRATE = 40
            .SPEEDMONITOR = False
            .RSMON_SPEED = True
            .RSMON_DISTANCE = False
            .SPEEDMON_REFRESHRATE = 20
            .STARTING_GRID_BEHIND = False
            .STARTING_GRID_IN = True
            .SPEECH = False
            .TIDE = 1
            .LUGWORMS = 1
            .WATER_SPLASHES = True
            .SAND_DUST = True
            .SPACE_PLANET = enumPlanets.moon
            .SPACE_ALIENS = enumAliens.unfriendly
            .SPACE_KIDNAPPINGRATE = 2
            .TUNNEL_COUNT = 2
            .TUNNEL_LENGTH = 300
            .TRIBUNES = True
            .SPECTATORS = 60
            .PARTICULATES_SLIDER = 3
            .COL_BACK = xlNone
            .COL_TEXT = vbBlack
            .COL_RANKINGS = vbWhite
            .DAYLIGHT = 0 'High noon (12:00)
            .AUTOFIT = True
            .AUTO_SAVE = False
            .ROADCROSSING = True
            .POLICECAR = True
        End With

        objRace.LOADED = False
        
    'Set the default path for GaloppSim files
        g_defaultPath = Application.UserLibraryPath 'Add-in folder
        
    'Determine the standard language according to the Excel version
        Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
            'German speaking countries
            Case msoLanguageIDGerman, msoLanguageIDGermanAustria, _
                    msoLanguageIDGermanLiechtenstein, msoLanguageIDGermanLuxembourg
                objOption.language = "DE"
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

    'Space on the top of the worksheet for control elements like the dropdown and menu buttons
    If g_strPlayMode = "RS" Then
        objBasicData.TOP_ROWS = 8
    Else 'AI edition
        objBasicData.TOP_ROWS = 0
    End If
    
    If g_strPlayMode = "RS" Then
        With Application
            .ActiveWindow.WindowState = xlMaximized 'Maximize the window
            objBasicData.XL_STYLE_NOTATION = .ReferenceStyle 'Save the current value
            .ReferenceStyle = xlA1 'Force A1-style notation
        End With
        Call RS_NewRace
    End If
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "Workbook_Open()")
    Call basAuxiliary.CodeCrash
End Sub

'Call the Super Developer Tools on double clicking the jockey´s eye
Private Sub Workbook_SheetBeforeDoubleClick(ByVal sh As Object, ByVal Target As Range, Cancel As Boolean)
    If sh.name = "GALOPPSIM" And Target.Address = Range("P11").Address Then
        frmSuperDev.show (vbModeless)
        Range("U15").Select
    End If
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
'        Application.ReferenceStyle = objOption.XL_STYLE_NOTATION 'Restore the setting
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
