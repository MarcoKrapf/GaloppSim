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

Public WithEvents RSButtonObject As MSForms.CommandButton
Attribute RSButtonObject.VB_VarHelpID = -1

'VARIABLES
Private nameID As String

'CLICK EVENT
Private Sub RSButtonObject_Click()
    basMainCode.RS_ExecuteClick (nameID)
End Sub

'PROPERTIES
Public Property Get RSbtnID() As String
    RSbtnID = nameID
End Property

Public Property Let RSbtnID(ByVal id As String)
    nameID = id
End Property
