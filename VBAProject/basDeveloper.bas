Attribute VB_Name = "basDeveloper"
Option Explicit
Option Private Module

'This module contains non-productive procedures
'which can be used for development purposes
'   Module basDeveloper

'Execute a procedure manually by placing the cursor
'inside and pressing F5

'Switch on/off tools and change parameters for development purposes
Sub DevelopmentTools()
        g_skipDelay = True 'Set true to skip delay commands (Application.Wait)
        g_payoutLogging = False 'Pay-out logging on(True)/off(False)
        g_errorLogPath = Environ("UserProfile") '& "\Desktop" 'Choose the path for the error log file
        g_errorLogging = True 'Error logging on(True)/off(False)
        g_defaultMLpath = Environ("UserProfile") 'Default path for Machine Learning export
        g_MLdataFileName = "GALOPPSIM_ML_DATA" 'File name for Machine Learning export
        g_defaultAutoSavePath = Environ("UserProfile") 'Default path for the auto-save function after a race
End Sub

'Execute this procedure to display the location of the error log file
Sub WhereIsTheErrorLog()
    Debug.Print g_errorLogPath & Application.PathSeparator & g_c_errorLogFileName & ".txt"
End Sub

'Execute this procedure to force an error for testing the error logging
Sub ForceError()
    On Error GoTo ERRORHANDLING
    Debug.Print 1 / 0
    Exit Sub
ERRORHANDLING:
    Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "ForceError()", "[raised for testing purposes]")
    Call CodeCrash
End Sub

'Execute this procedure to open the error log file in the Notepad (Microsoft Windows)
Sub OpenErrorLog()
    Dim errorfile As String
    errorfile = g_errorLogPath & Application.PathSeparator & g_c_errorLogFileName & ".txt"
    Shell "cmd /c " & Chr(34) & errorfile & Chr(34) 'Chr(34) represents the double quotes "
End Sub

'Execute this procedure for writing the value of the background colour
'of each selected cell into this cell while designing a new race
Sub ReadColourValues()
    Dim r As Range
    For Each r In Selection
        r.Value = r.Interior.Color
    Next
End Sub

'Execute this procedure for testing the generation of random integer values
'with the formula Int((upperBound - lowerBound + 1) * Rnd + lowerBound)
Sub RandomVerification()
    Dim loops As Long, loopNr As Long, _
        upperBound As Long, lowerBound As Long, _
        current As Long, min As Long, max As Long
    'Modify the values if needed
    loops = 10000 '10 'Number of random values to be generated
    upperBound = 142 '16777215 'Maximum value
    lowerBound = 2 '0 'Minimum value
    'Set the start values
    max = lowerBound
    min = upperBound
    'Generate random integer numbers
    For loopNr = 1 To loops
        Randomize 'Initialize the random-number generator (create a new seed value)
        current = Int((upperBound - lowerBound + 1) * Rnd + lowerBound)
        If current < min Then min = current
        If current > max Then max = current
    Next
    'Evaluate the results
    Debug.Print
    Debug.Print "Random values generated: " & loopNr - 1
    Debug.Print "Highest value generated: " & max & " (upper bound is " & upperBound & ")"
    Debug.Print "Lowest  value generated: " & min & " (lower bound is " & lowerBound & ")"
End Sub

'Use this procedure for initializing a new pay-out log file
Sub NewPayoutLogFile()
    Open Environ("UserProfile") & "\Desktop\GALOPPSIM_PAYOUTLOG.csv" For Output As #1
        Print #1, "Date" & ";" & "Level" & ";" & "Race ID" & ";" _
        & "Running horses" & ";" & "Bet slips" & ";" _
        & "Type of bet" & ";" & "Stake (EUR)" & ";" & "Pay-out (EUR)"
    Close #1
End Sub


'Fill in the brackets a data type and execute this procedure for getting the name of the type
Sub GetTypeName()
    Debug.Print TypeName(xlNone) 'TypeName(...) e.g. xlSolid --> type Long
End Sub

