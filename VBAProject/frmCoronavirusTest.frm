VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCoronavirusTest 
   ClientHeight    =   7824
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10668
   OleObjectBlob   =   "frmCoronavirusTest.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmCoronavirusTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Pop-up with the Coronavirus test centre
'   UserForm frmCoronavirusTest

Dim arr_intTestOrder() As Integer 'Order for test execution
Dim arr_objLabelHorse() As Variant 'MSForms.Label
Dim arr_objLabelStatus() As Variant 'MSForms.Label
Dim i As Integer, j As Integer

Const positivrate As Integer = 20 '(%)
Const protectionrate As Integer = 95 '(%)
Const virusmutations As Integer = 5 '(% of all are infected)

Private Sub UserForm_Initialize()
    Dim X As Integer, Y As Integer
    
    ReDim arr_objLabelHorse(1 To objRace.NUMBER_ENROLLED)
    ReDim arr_objLabelStatus(1 To objRace.NUMBER_ENROLLED)
    
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    With lblCaption
        .TextAlign = fmTextAlignCenter
        .caption = GetText(g_arr_Text, "CORONA001")
    End With
    
    X = objRace.NUMBER_STARTING
    
    For i = 1 To objRace.NUMBER_ENROLLED
    
        'Create a label with the horse number and name
        Set arr_objLabelHorse(i) = Me.Controls.Add("Forms.Label.1", , True)
        With arr_objLabelHorse(i)
            .name = "#" & g_arr_varHorses(i, 11)
            With .Font
                .name = "Tahoma"
                .size = 12
            End With
            .left = 12
            .top = 46 + ((i - 1) * 18)
            .width = 250
            .Height = 16
            .TextAlign = fmTextAlignLeft
            .caption = "#" & g_arr_varHorses(i, 11) & " " & g_arr_varHorses(i, 1)
            If g_arr_varHorses(i, 0) = "CANCELLED" Then .Font.Strikethrough = True
        End With
        
        'Create a label with the coronavirus status
        Set arr_objLabelStatus(i) = Me.Controls.Add("Forms.Label.1", , True)
        With arr_objLabelStatus(i)
            .name = g_arr_varHorses(i, 11)
            With .Font
                .name = "Tahoma"
                .size = 12
            End With
            .left = 280
            .top = 46 + ((i - 1) * 18)
            .width = 250
            .Height = 16
            If g_arr_varHorses(i, 0) = "START" Then
                If g_arr_varHorses(i, 24) = "VACCINATED" Then
                    .caption = GetText(g_arr_Text, "CORONA008")
                ElseIf g_arr_varHorses(i, 24) = "NOT VACCINATED" Then
                    .caption = GetText(g_arr_Text, "CORONA009")
                Else
                    .caption = ""
                End If
            ElseIf g_arr_varHorses(i, 0) = "CANCELLED" Then
                .caption = GetText(g_arr_Text, "CORONA010")
            Else
                .caption = g_arr_varHorses(i, 0)
                .BackColor = RGB(200, 200, 200) 'Grey
            End If
            .TextAlign = fmTextAlignLeft
            .AutoSize = True
        End With
    Next i
    
    'Align the test start button
    With btnStartTest
        With .Font
            .name = "Tahoma"
            .size = 12
        End With
        .left = 12
        .top = 60 + (objRace.NUMBER_ENROLLED * 18)
        .width = 120
        .Height = 48
        .Font.Bold = True
        .WordWrap = True
        .caption = GetText(g_arr_Text, "CORONA002")
    End With
    
    'Align the leave button
    With btnFinished
        With .Font
            .name = "Tahoma"
            .size = 12
        End With
        .left = 410
        .top = 60 + (objRace.NUMBER_ENROLLED * 18)
        .width = 120
        .Height = 48
        .Font.Bold = True
        .WordWrap = True
        .Visible = False
        .caption = GetText(g_arr_Text, "CORONA003")
    End With
    
    'Aling the labels for the test progress bar
    With lblProgressBar1
        .left = 140
        .top = 60 + (objRace.NUMBER_ENROLLED * 18)
        .width = 260
        .Height = 48
        .BackColor = vbBlack 'vbWhite
        .caption = ""
    End With
    With lblProgressBar2
        .left = 140
        .top = 60 + (objRace.NUMBER_ENROLLED * 18)
        .width = 0
        .Height = 48
        .BackColor = vbBlack
        .caption = ""
    End With
    
    'Adjust the UserForm
    With Me
        .caption = objRace.RACE_NAME & " " & objRace.RACE_YEAR 'Race name and year
        .width = 550
        .Height = 145 + 18 * objRace.NUMBER_ENROLLED
    End With
    
    'Determine the test execution order
    ReDim arr_intTestOrder(1 To X)
    For i = 1 To objRace.NUMBER_ENROLLED
        If g_arr_varHorses(i, 0) = "START" Then  'Needs to be tested
            Do 'Find an empty position in the array
                Y = Int((X - 1 + 1) * Rnd + 1)
                If arr_intTestOrder(Y) = 0 Then
                    arr_intTestOrder(Y) = g_arr_varHorses(i, 11)
                    Exit Do
                End If
            Loop
        End If
    Next i
    
    'Display the UserForm in the center of the Window
    Call basAuxiliary.PlaceUserFormInCenter(Me)
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "UserForm_Initialize()")
    Call basAuxiliary.CodeCrash
    
End Sub

'Click on the "Start testing" button
Private Sub btnStartTest_Click()
    Dim progress As Double, loops As Double
    Dim m As Integer, n As Integer
    
    On Error GoTo ERRORHANDLING 'In case an error occurs
    
    btnStartTest.caption = GetText(g_arr_Text, "CORONA007")
    
    'Testing starts
    loops = 20000 'Number of loops
    m = 1
    
    #If Debugging Then
        Debug.Print "CORONAVIRUS TEST CENTRE"
        Debug.Print
        Debug.Print "Infection rate         : " & positivrate & "%"
        Debug.Print "Vaccination protection : " & protectionrate & "%"
    #End If
    
    For progress = 0 To loops
        If progress = Round(((loops / 2) + (loops / 2) / UBound(arr_intTestOrder()) * m), 0) Then

            If g_arr_varHorses(arr_intTestOrder(m), 24) = "VACCINATED" Then
                Randomize
                i = Int((((100 / positivrate) - 1) - 0 + 1) * Rnd + 0)
                Randomize
                j = Int((((100 / (100 - protectionrate)) - 1) - 0 + 1) * Rnd + 0) 'Vaccination
            ElseIf g_arr_varHorses(arr_intTestOrder(m), 24) = "NOT VACCINATED" Then
                Randomize
                i = Int((((100 / positivrate) - 1) - 0 + 1) * Rnd + 0)
                j = 0 'No vaccination
            Else
                Randomize
                i = Int(((100 / virusmutations) - 0 + 1) * Rnd + 0) 'Virus mutations
                j = 0
            End If
            
            #If Debugging Then
                Debug.Print vbNewLine & ">> TESTING NR #" & arr_intTestOrder(m) & " " & g_arr_varHorses(arr_intTestOrder(m), 1)
                If g_arr_varHorses(arr_intTestOrder(m), 24) = "VACCINATED" Then Debug.Print "   (vaccinated)"
                Debug.Print "   Infection rate         : " & i
                Debug.Print "   Vaccination protection : " & j
            #End If
            
            'Check wheter the test is positive or negative
            If i + j = 0 Then
                #If Debugging Then
                    Debug.Print "   ----- INFECTED -----"
                #End If
                g_arr_varHorses(arr_intTestOrder(m), 0) = "CORONAVIRUSPOSITIVE"
                objRace.NUMBER_STARTING = objRace.NUMBER_STARTING - 1
                With arr_objLabelStatus(arr_intTestOrder(m))
                    .width = 250
                    .caption = GetText(g_arr_Text, "CORONA005")
                    .AutoSize = True
                    .BackColor = vbRed
                End With
            End If
            m = m + 1
        End If
        
        'Change the colour of the progress bar
'        lblProgressBar1.BackColor = Int((16777215 - 0 + 1) * Rnd + 0)
        With lblProgressBar2
            .width = lblProgressBar1.width / loops * progress
            .BackColor = Int((16777215 - 0 + 1) * Rnd + 0)
        End With
        
        DoEvents
    Next progress
    
    'Testing finished
    
    btnStartTest.caption = GetText(g_arr_Text, "CORONA004")
    If Not g_skipDelay Then Application.Wait (Now + TimeValue("0:00:04"))
    
    lblProgressBar1.Visible = False
    lblProgressBar2.Visible = False
    
    Dim ctr As control
    For Each ctr In Me.Controls
        If IsNumeric(ctr.name) Then
            If g_arr_varHorses(ctr.name, 0) = "START" Then
                With arr_objLabelStatus(ctr.name)
                    .width = 250
                    .caption = GetText(g_arr_Text, "CORONA006")
                    .AutoSize = True
                End With
            End If
        End If
    Next
    
    btnStartTest.Visible = False
    btnFinished.Visible = True
    
    Exit Sub
ERRORHANDLING:
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "btnStartTest_Click()")
    Call basAuxiliary.CodeCrash
    
End Sub

'Click on the "Leave testing are" button
Private Sub btnFinished_Click()
    Unload frmCoronavirusTest 'Close the popup
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then 'Click on "X" in the upper right corner of the UserForm
        'Pop-up with a security check before cancelling the race
        Call ShowMessagePopup(objRace.RACE_NAME & " " & objRace.RACE_YEAR, GetText(g_arr_Text, "CORONA011"), _
            enumButton.YesNo, vbModal)
        'Evaluate the return value
        If g_enumButton = enumButton.yes Then
            Unload frmCoronavirusTest ''>>Yes<< clicked: close the popup
        Else '>>No<< clicked
            Cancel = 1 'Don´t close the pop-up
        End If
    End If
End Sub

