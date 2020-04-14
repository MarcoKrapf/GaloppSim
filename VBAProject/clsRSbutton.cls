VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRSbutton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Class Module for creating object instances for each button on the GALOPPSIM Worksheet
'   Class Module clsRSbutton

'Enable the class for handling events of Command Buttons on the Worksheet
Public WithEvents RSButtonObject As MSForms.CommandButton
Attribute RSButtonObject.VB_VarHelpID = -1

'OBJECT VARIABLE
Private nameID As String

'PROPERTIES
Public Property Get RSbtnID() As String
    RSbtnID = nameID
End Property

Public Property Let RSbtnID(ByVal id As String)
    nameID = id
End Property

'BUTTON CLICK EVENT
Private Sub RSButtonObject_Click()
    basMainCode.RS_ExecuteClick (nameID) 'Pass the ID of the button
End Sub

Private Sub Class_Terminate()
    If g_errorLogging Then Call WriteErrorLog(VBA.Now, Err, Application.VBE.ActiveCodePane.CodeModule, "Class_Terminate()", "RSbtnID: " & Me.RSbtnID)
    
    'Show a classical MessageBox
    MsgBox "Instances of the Class Module 'clsRSbutton' terminated unexpectedly." _
        & vbNewLine & "Please close and restart the Workbook.", vbExclamation, "ATTENTION"
End Sub
